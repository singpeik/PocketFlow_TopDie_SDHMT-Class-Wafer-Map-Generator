import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.colors import to_rgba
from matplotlib.markers import MarkerStyle
from pptx import Presentation
from openpyxl import load_workbook

from flow import create_ioe_flow
from nodes import CollectIOEInputsNode, GenerateIOEMRBReportNode, GenerateIOEWaferMapNode
from utils.ioe_workflow import (
    VID_TO_IOE_WLA_SHEET,
    _replace_excel_sheet,
    build_ioe_wafer_map,
    normalize_lot7_wid3,
    normalize_lot7_wid3_list,
)
from utils.reporting import (
    CLASS_FAIL_COLOR,
    SDHMT_FAIL_COLOR,
    _build_step6_summary_row,
    _combined_style,
    _draw_ioe_combined_map,
    _draw_ioe_map,
    _find_edge_dies,
    _interest_bin_marker,
    _interest_failure_font_style,
    _is_interest_fail,
    _class_map_style,
    _sdhmt_map_style,
    generate_ioe_mrb_summary_ppt,
    write_wafer_result_summary,
)


class InterestBinTests(unittest.TestCase):
    def setUp(self):
        self.geometry = [(x, y) for x in range(3) for y in range(3)]
        self.rows = [
            {'FAB_LOT': 'TESTLOT', 'FAB_WAFER': 1, 'X_LOCATION': 0, 'Y_LOCATION': 1,
             'SDHMT IB': 15, 'SDHMT FB': 1590, 'Class IB': 1, 'Class FB': 100},
            {'FAB_LOT': 'TESTLOT', 'FAB_WAFER': 1, 'X_LOCATION': 1, 'Y_LOCATION': 1,
             'SDHMT IB': 36, 'Class IB': 15, 'Class FB': 1599},
            {'FAB_LOT': 'TESTLOT', 'FAB_WAFER': 1, 'INTEL_DIE_INDEX': 'wafer_2_1',
             'SDHMT IB': 1, 'Class IB': 36},
        ]

    def test_bin_boundaries(self):
        cases = [(15, 1589, True), (15, 1590, True), (15, 1599, True),
             (15, 1600, True), (36, None, True), (36, 3600, True),
             (1, 1594, False), (None, 1594, False), (15, None, True),
                 ('15', '1594', True)]
        for test_name in ('SDHMT', 'Class'):
            for ib_value, fb_value, expected in cases:
                with self.subTest(test=test_name, ib=ib_value, fb=fb_value):
                    row = {f'{test_name} IB': ib_value, f'{test_name} FB': fb_value}
                    self.assertEqual(_is_interest_fail(row, test_name), expected)

    def test_product_neutral_labels(self):
        self.assertEqual(VID_TO_IOE_WLA_SHEET, 'VID to Wafer&WLA')
        for normalize, value in ((normalize_lot7_wid3, ''),
                                 (normalize_lot7_wid3, 'invalid'),
                                 (normalize_lot7_wid3_list, '')):
            with self.assertRaises(ValueError) as error:
                normalize(value)
            self.assertNotIn('IOE', str(error.exception))
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / 'Final.xlsx'
            pd.DataFrame(self.rows).to_excel(workbook, sheet_name='CATTS-cleaned', index=False)
            _replace_excel_sheet(workbook, VID_TO_IOE_WLA_SHEET, pd.DataFrame(self.rows))
            with pd.ExcelFile(workbook) as excel:
                self.assertIn('VID to Wafer&WLA', excel.sheet_names)
                self.assertNotIn('VID to IOE&WLA', excel.sheet_names)

    def test_failure_font_emphasis(self):
        for column, value, color in (
            ('SDHMT Whole Fails', '2', '#C00000'),
            ('Class Edge Fails', '1', '#C65D00'),
            ('Combined Edge Fail?', 'Yes', '#C65D00'),
            ('SDHMT IB15 Other FB', '3', '#C00000'),
            ('Class IB36', '1', '#C00000'),
        ):
            with self.subTest(column=column):
                self.assertEqual(_interest_failure_font_style({}, column, value),
                                 {'bold': True, 'font_color': color})
        for column, value in (('SDHMT Whole Fails', '0'), ('Class Edge Fails', '0'),
                              ('Combined Edge Fail?', 'No'), ('Class Edge Fail?', 'Unknown'),
                              ('Total Die Tested', '100'), ('Test %', '70.0%')):
            self.assertIsNone(_interest_failure_font_style({}, column, value))

    def test_edge_geometry_and_union(self):
        edge_dies = _find_edge_dies(self.geometry)
        self.assertEqual(len(edge_dies), 9)
        self.assertIn((1, 1), edge_dies)
        self.assertEqual(len(_find_edge_dies(self.geometry, layers=1)), 8)
        summary = _build_step6_summary_row('TESTLOT_001', pd.DataFrame(self.rows), self.geometry)
        for test_name in ('SDHMT', 'Class'):
            self.assertEqual(summary[f'{test_name} Whole Fails'], 2)
            self.assertEqual(summary[f'{test_name} Edge Fails'], 2)
        self.assertEqual(summary['Combined Whole Fails'], 3)
        self.assertEqual(summary['Combined Edge Fails'], 3)
        self.assertEqual(summary['Combined Edge Fail?'], 'Yes')
        interior = _build_step6_summary_row('TESTLOT_001', pd.DataFrame([self.rows[1]]),
                            [(x, y) for x in range(-2, 5) for y in range(-2, 5)])
        self.assertEqual(interior['Combined Edge Fail?'], 'No')

    def test_combined_map_preserves_either_failure(self):
        self.assertEqual(_combined_style(self.rows[0]), SDHMT_FAIL_COLOR)
        self.assertEqual(_combined_style(self.rows[1]), CLASS_FAIL_COLOR)
        self.assertEqual(_combined_style(self.rows[2]), CLASS_FAIL_COLOR)

    def test_outer_edge_excludes_enclosed_holes(self):
        full_geometry = {(x, y) for x in range(-3, 4) for y in range(-3, 4)}
        expected = {(x, y) for x, y in full_geometry if abs(x) >= 2 or abs(y) >= 2}
        geometry = full_geometry - {(0, 0), (1, 0)}
        self.assertEqual(_find_edge_dies(geometry), expected)
        hole_neighbor = pd.DataFrame([{'X_LOCATION': -1, 'Y_LOCATION': 0, 'SDHMT IB': 36}])
        summary = _build_step6_summary_row('TESTLOT_001', hole_neighbor, geometry)
        self.assertEqual(summary['SDHMT Edge Fails'], 0)
        self.assertEqual(summary['SDHMT Edge Fail?'], 'No')

    def test_outer_edge_includes_exterior_notches(self):
        full_geometry = {(x, y) for x in range(7) for y in range(7)}
        geometry = full_geometry - {(3, 0), (3, 1), (3, 2)}
        edges = _find_edge_dies(geometry)
        self.assertIn((3, 3), edges)
        self.assertIn((2, 1), edges)
        self.assertNotIn((4, 4), edges)

    def test_two_layer_edge_band(self):
        geometry = {(x, y) for x in range(9) for y in range(9)} - {(4, 4)}
        outer = _find_edge_dies(geometry, layers=1)
        band = _find_edge_dies(geometry)
        self.assertEqual(len(outer), 32)
        self.assertEqual(len(band), 56)
        self.assertTrue(outer < band <= geometry)
        self.assertIn((1, 4), band)
        self.assertNotIn((2, 4), band)
        self.assertNotIn((3, 4), band)
        rows = [{'X_LOCATION': 1, 'Y_LOCATION': 4, 'SDHMT IB': 36},
                {'X_LOCATION': 2, 'Y_LOCATION': 4, 'SDHMT IB': 36}]
        summary = _build_step6_summary_row('TESTLOT_001', pd.DataFrame(rows), geometry)
        self.assertEqual(summary['SDHMT Edge Fails'], 1)
        for invalid in (0, -1, 1.5, True):
            with self.assertRaises(ValueError):
                _find_edge_dies(geometry, layers=invalid)

    def test_tested_die_counts(self):
        rows = [
            {'X_LOCATION': 0, 'Y_LOCATION': 0, 'SDHMT IB': 1, 'Class IB': 1},
            {'X_LOCATION': 0, 'Y_LOCATION': 0, 'SDHMT IB': 1, 'Class IB': 1},
            {'X_LOCATION': 0, 'Y_LOCATION': 1, 'SDHMT IB': 36},
            {'X_LOCATION': 1, 'Y_LOCATION': 0, 'SDHMT IB': 6, 'Class IB': 36},
            {'X_LOCATION': 1, 'Y_LOCATION': 1, 'SDHMT IB': 6},
            {'X_LOCATION': 2, 'Y_LOCATION': 2},
        ]
        summary = _build_step6_summary_row('TESTLOT_001', pd.DataFrame(rows), self.geometry)
        self.assertEqual(summary['SDHMT Tested Dies'], 2)
        self.assertEqual(summary['Class Tested Dies'], 2)
        self.assertEqual(summary['Combined Tested Dies'], 3)

    def test_test_percentage_uses_selected_map(self):
        frame = pd.DataFrame(self.rows + [self.rows[0]])
        for geometry in (self.geometry, self.geometry + [self.geometry[0]],
                         [(x, y) for x in range(4) for y in range(4)]):
            with self.subTest(map_dies=len(set(geometry))):
                summary = _build_step6_summary_row('TESTLOT_001', frame, geometry)
                self.assertEqual(summary['Wafer Map Dies'], len(set(geometry)))
                self.assertEqual(summary['Total Die Tested'], 3)
                self.assertAlmostEqual(summary['Test %'], 3 / len(set(geometry)))
        untested = pd.DataFrame([{'X_LOCATION': 0, 'Y_LOCATION': 0, 'SDHMT IB': 6}])
        summary = _build_step6_summary_row('TESTLOT_001', untested, self.geometry)
        self.assertEqual(summary['Total Die Tested'], 0)
        self.assertEqual(summary['Test %'], 0)

    def test_tested_die_counts_without_coordinates(self):
        rows = [{'INTEL_DIE_INDEX': 'missing-coordinate', 'SDHMT IB': 1},
                {'INTEL_DIE_INDEX': 'missing-coordinate', 'SDHMT IB': 1},
                {'VISUAL_ID': 'VID123', 'Class IB': 1},
                {'VISUAL_ID': 'VID123', 'Class IB': 1}]
        summary = _build_step6_summary_row('TESTLOT_001', pd.DataFrame(rows), self.geometry)
        self.assertEqual(summary['SDHMT Tested Dies'], 1)
        self.assertEqual(summary['Class Tested Dies'], 1)
        self.assertEqual(summary['Combined Tested Dies'], 2)

    def test_unknown_coordinates_and_missing_geometry(self):
        for row in ({'SDHMT IB': 36}, {'SDHMT IB': 36, 'X_LOCATION': 99, 'Y_LOCATION': 99}):
            summary = _build_step6_summary_row('TESTLOT_001', pd.DataFrame([row]), self.geometry)
            self.assertEqual(summary['Combined Edge Fail?'], 'Unknown')
        with self.assertRaises(ValueError):
            _find_edge_dies([])

    def test_interest_bin_marker_shapes(self):
        for test_name in ('SDHMT', 'Class'):
            for fb_value in range(1590, 1600):
                row = {f'{test_name} IB': 15, f'{test_name} FB': fb_value}
                self.assertEqual(_interest_bin_marker(row, test_name), '^')
            self.assertEqual(_interest_bin_marker({f'{test_name} IB': 36}, test_name), 'D')
            self.assertEqual(_interest_bin_marker({f'{test_name} IB': 15, f'{test_name} FB': 1589}, test_name), 'v')
            self.assertEqual(_interest_bin_marker({f'{test_name} IB': 15}, test_name), 'v')

    def test_all_ib15_category_counts_and_edges(self):
        rows = [
            {'X_LOCATION': 0, 'Y_LOCATION': 0, 'SDHMT IB': 15, 'SDHMT FB': 1589,
             'Class IB': 15, 'Class FB': 1590},
            {'X_LOCATION': 1, 'Y_LOCATION': 1, 'SDHMT IB': 15, 'SDHMT FB': 1599,
             'Class IB': 15, 'Class FB': 1600},
            {'X_LOCATION': 2, 'Y_LOCATION': 2, 'SDHMT IB': 15, 'Class IB': 15},
            {'X_LOCATION': 0, 'Y_LOCATION': 2, 'SDHMT IB': 36, 'Class IB': 36},
        ]
        summary = _build_step6_summary_row('TESTLOT_001', pd.DataFrame(rows), self.geometry)
        for test_name in ('SDHMT', 'Class'):
            self.assertEqual(summary[f'{test_name} Whole Fails'], 4)
            self.assertEqual(summary[f'{test_name} Edge Fails'], 4)
            for category in ('IB15 FB1590-1599', 'IB15 Other FB', 'IB15 Missing FB', 'IB36'):
                self.assertEqual(summary[f'{test_name} {category}'], 1)
        self.assertEqual(summary['Combined Whole Fails'], 4)
        self.assertEqual(summary['Combined Edge Fails'], 4)
        self.assertEqual(summary['Combined IB15 FB1590-1599'], 2)
        self.assertEqual(summary['Combined IB15 Other FB'], 2)
        self.assertEqual(summary['Combined IB15 Missing FB'], 1)
        self.assertEqual(summary['Combined IB36'], 1)
        with patch('utils.reporting.plt.savefig'), patch('utils.reporting.plt.close'):
            _draw_ioe_map('unused.png', self.geometry, pd.DataFrame(rows), 'All IB15', _sdhmt_map_style)
            plotted = plt.gcf().axes[0].collections[1:]
            marker = MarkerStyle('v')
            expected = marker.get_path().transformed(marker.get_transform()).vertices.tolist()
            self.assertEqual(plotted[0].get_paths()[0].vertices.tolist(), expected)
            self.assertEqual(plotted[2].get_paths()[0].vertices.tolist(), expected)
        plt.close('all')

    def test_plotted_shapes_and_combined_overlap(self):
        close_figure = plt.close
        with patch('utils.reporting.plt.savefig'), patch('utils.reporting.plt.close'):
            for test_name, resolver in (('SDHMT', _sdhmt_map_style), ('Class', _class_map_style)):
                _draw_ioe_map('unused.png', self.geometry, pd.DataFrame(self.rows), 'Test', resolver, test_name)
                figure = plt.gcf()
                markers = [collection.get_paths()[0].vertices.tolist() for collection in figure.axes[0].collections[1:]]
                for expected in ('^', 'D'):
                    style = MarkerStyle(expected)
                    self.assertIn(style.get_path().transformed(style.get_transform()).vertices.tolist(), markers)
                close_figure(figure)
            _draw_ioe_combined_map('unused.png', self.geometry, pd.DataFrame([self.rows[1]]), 'Both tests')
            collections = plt.gcf().axes[0].collections[1:]
            self.assertEqual(len(collections), 2)
            self.assertEqual(collections[0].get_facecolors()[0][:3].tolist(), list(to_rgba(SDHMT_FAIL_COLOR)[:3]))
            self.assertEqual(collections[1].get_facecolors()[0][:3].tolist(), list(to_rgba(CLASS_FAIL_COLOR)[:3]))
            self.assertEqual(collections[0].get_offsets().tolist(), collections[1].get_offsets().tolist())
            self.assertGreater(collections[0].get_sizes()[0], collections[1].get_sizes()[0])
        plt.close('all')

    def test_gpu_geometry_conversion_preserves_ioe(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            source = folder / 'GPU.xlsx'
            pd.DataFrame({'FAB_WAFER': [902, 905, 905], 'X_LOCATION': [0, 1, 1],
                          'Y_LOCATION': [0, 0, 0]}).to_excel(source, index=False)
            library = folder / 'maps.json'
            ioe_entry = {'active_dies': [[10, 10]]}
            with library.open('w', encoding='utf-8') as stream:
                json.dump({'IOE': ioe_entry}, stream)
            output = folder / 'gpu.json'
            payload = build_ioe_wafer_map(source, output, library, map_name='GPU')
            self.assertEqual(payload['map_name'], 'GPU')
            self.assertEqual(payload['active_die_count'], 2)
            self.assertEqual(payload['wafer_count'], 2)
            self.assertEqual(payload['active_dies'], [[0, 0], [1, 0]])
            maps = json.loads(library.read_text(encoding='utf-8'))
            self.assertEqual(maps['IOE'], ioe_entry)
            self.assertEqual(maps['GPU']['active_dies'], payload['active_dies'])
            cached = build_ioe_wafer_map(folder / 'missing.xlsx', output, library, map_name='GPU')
            self.assertEqual(cached['build_mode'], 'reused_json')
            with self.assertRaises(ValueError):
                build_ioe_wafer_map(source, output, library, map_name='IOE')

    def test_invalid_geometry_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            for x_value in (None, 0.5):
                with self.subTest(x=x_value):
                    source = folder / 'GPU.xlsx'
                    pd.DataFrame({'X_LOCATION': [x_value], 'Y_LOCATION': [0]}).to_excel(source, index=False)
                    with self.assertRaises(ValueError):
                        build_ioe_wafer_map(source, folder / 'gpu.json', folder / 'maps.json', map_name='GPU')

    def test_product_geometry_routing(self):
        for stop_after_step in (4, 5, 6):
            self.assertIsInstance(create_ioe_flow(stop_after_step).start_node, CollectIOEInputsNode)
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            workbook = folder / 'Final.xlsx'
            pd.DataFrame(self.rows).to_excel(workbook, sheet_name='CATTS-cleaned', index=False)
            for product in ('IOE', 'GPU'):
                with self.subTest(product=product):
                    node = GenerateIOEWaferMapNode()
                    path = folder / f'{product.lower()}.json'
                    node.set_params({f'{product.lower()}_map_output': str(path)})
                    shared = {'selected_wafer': {'product': product}}
                    prep = node.prep(shared)
                    self.assertEqual(prep['map_name'], product)
                    node.post(shared, prep, {'map_name': product})
                    self.assertEqual(Path(shared['wafer_map_path']), path)
                    shared['filtered_catts_raw_summary'] = {'final_workbook': str(workbook)}
                    report_node = GenerateIOEMRBReportNode()
                    report_node.set_params({'output_dir': str(folder)})
                    report_prep = report_node.prep(shared)
                    self.assertEqual(Path(report_prep['ioe_map_path']), path)
                    self.assertIn(f'{product} MRB', Path(report_prep['mrb_ppt_output']).name)

    def test_excel_summary_table_and_refresh(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'Final.xlsx'
            pd.DataFrame(self.rows).to_excel(path, sheet_name='CATTS-cleaned', index=False)
            workbook = load_workbook(path)
            workbook.create_sheet('Notes')['A1'] = '=1+1'
            workbook.save(path)
            workbook.close()
            summary = _build_step6_summary_row('TESTLOT_001', pd.DataFrame(self.rows), self.geometry)
            second = {**summary, 'LOT7_WID3': 'TESTLOT_002'}
            self.assertEqual(write_wafer_result_summary(path, [summary, second]), 'Result Summary')
            workbook = load_workbook(path)
            try:
                self.assertEqual(workbook['Notes']['A1'].value, '=1+1')
                self.assertEqual(workbook['CATTS-cleaned'].max_row, len(self.rows) + 1)
                sheet = workbook['Result Summary']
                headers = [cell.value for cell in sheet[1]]
                self.assertEqual(headers[:4], ['LOT7_WID3', 'Wafer Map Dies', 'Total Die Tested', 'Test %'])
                self.assertEqual(sheet.freeze_panes, 'B2')
                self.assertEqual(sheet.max_row, 3)
                table = next(iter(sheet.tables.values()))
                self.assertEqual(table.ref, sheet.dimensions)
                self.assertEqual(table.autoFilter.ref, sheet.dimensions)
                self.assertAlmostEqual(sheet['D2'].value, 1 / 3)
                self.assertEqual(sheet['D2'].number_format, '0.0%')
                self.assertIn('Class IB15 Other FB', headers)
                for column, color in (('SDHMT Whole Fails', '00C00000'),
                                      ('SDHMT Edge Fails', '00C65D00')):
                    cell = sheet.cell(2, headers.index(column) + 1)
                    self.assertTrue(cell.font.bold)
                    self.assertEqual(cell.font.color.rgb, color)
                self.assertFalse(sheet['C2'].font.bold)
            finally:
                workbook.close()
            write_wafer_result_summary(path, [second])
            workbook = load_workbook(path)
            try:
                self.assertEqual(workbook.sheetnames.count('Result Summary'), 1)
                self.assertEqual(workbook['Result Summary'].max_row, 2)
                self.assertEqual(workbook['Result Summary']['A2'].value, 'TESTLOT_002')
                self.assertEqual(workbook['Notes']['A1'].value, '=1+1')
            finally:
                workbook.close()

    def test_excel_summary_failure_preserves_workbook(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'Final.xlsx'
            pd.DataFrame(self.rows).to_excel(path, sheet_name='CATTS-cleaned', index=False)
            original = path.read_bytes()
            summary = _build_step6_summary_row('TESTLOT_001', pd.DataFrame(self.rows), self.geometry)
            with patch('utils.reporting.os.replace', side_effect=PermissionError('Workbook locked')):
                with self.assertRaises(PermissionError):
                    write_wafer_result_summary(path, [summary])
            self.assertEqual(path.read_bytes(), original)
            self.assertEqual(list(Path(directory).glob('*.xlsx')), [path])
            with self.assertRaises(ValueError):
                write_wafer_result_summary(path, [])

    def test_gpu_ppt_output(self):
        self._check_ppt_output('GPU')

    def test_ppt_output(self):
        self._check_ppt_output('IOE')

    def _check_ppt_output(self, product):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            workbook = folder / 'Final.xlsx'
            pd.DataFrame(self.rows).to_excel(workbook, sheet_name='CATTS-cleaned', index=False)
            map_path = folder / 'geometry.json'
            with map_path.open('w', encoding='utf-8') as stream:
                json.dump({'map_name': product, 'active_dies': self.geometry}, stream)
            result = generate_ioe_mrb_summary_ppt(workbook, map_path, folder / 'summary.pptx', folder)
            self.assertEqual(result['wafer_count'], 1)
            self.assertEqual(result['mdk_outputs'], [])
            self.assertEqual(result['summary_sheet'], 'Result Summary')
            excel_summary = pd.read_excel(workbook, sheet_name='Result Summary')
            self.assertEqual(excel_summary.iloc[0]['Total Die Tested'], 3)
            self.assertAlmostEqual(excel_summary.iloc[0]['Test %'], 1 / 3)
            self.assertEqual(excel_summary.iloc[0]['Combined Edge Fails'],
                             result['summary_rows'][0]['Combined Edge Fails'])
            self.assertFalse(list(folder.glob('*.csv')))
            self.assertFalse(list(folder.glob('*.png')))
            presentation = Presentation(result['output_ppt'])
            self.assertEqual(len(presentation.slides), 5)
            for slide in presentation.slides:
                self.assertTrue(slide.shapes.title.text.startswith(product))
            slide_text = '\n'.join(
                shape.text for slide in presentation.slides for shape in slide.shapes
                if shape.has_text_frame
            )
            table_text = '\n'.join(
                cell.text for slide in presentation.slides for shape in slide.shapes
                if shape.has_table for row in shape.table.rows for cell in row.cells
            )
            self.assertIn('edge fail? Yes', slide_text)
            self.assertIn('wafer test result summary', slide_text)
            self.assertIn('PocketFlow Author: Yang1.Xiang/ SingPei', slide_text)
            self.assertIn('Interest: all IB15 and IB36 results', slide_text)
            self.assertIn('IB15 other FB:', slide_text)
            self.assertIn('IB15 missing FB:', slide_text)
            self.assertIn('Combined Edge Fail?', table_text)
            for test_name in ('SDHMT', 'Class'):
                self.assertIn(f'{test_name} Tested Dies', table_text)
            self.assertIn('Wafer Map Dies', table_text)
            self.assertIn('Total Die Tested', table_text)
            self.assertIn('Test %', table_text)
            self.assertIn('33.3%', table_text)
            self.assertNotIn('Combined Tested Dies', table_text)
            table = next(shape.table for shape in presentation.slides[1].shapes if shape.has_table)
            headers = [cell.text for cell in table.rows[0].cells]
            for name, color in (('SDHMT Whole Fails', 'C00000'),
                                ('SDHMT Edge Fails', 'C65D00'),
                                ('SDHMT Edge Fail?', 'C65D00')):
                run = table.cell(1, headers.index(name)).text_frame.paragraphs[0].runs[0]
                self.assertTrue(run.font.bold)
                self.assertEqual(str(run.font.color.rgb), color)
            neutral = table.cell(1, headers.index('Total Die Tested')).text_frame.paragraphs[0].runs[0]
            self.assertFalse(neutral.font.bold)
            caption = next(shape for shape in presentation.slides[2].shapes
                           if shape.has_text_frame and shape.text.startswith('SDHMT tested dies:'))
            runs = caption.text_frame.paragraphs[0].runs
            self.assertFalse(runs[0].font.bold)
            self.assertEqual(str(runs[1].font.color.rgb), 'C00000')
            self.assertTrue(runs[1].font.bold)
            self.assertEqual(str(runs[2].font.color.rgb), 'C65D00')
            self.assertTrue(runs[2].font.bold)
            self.assertIn('SDHMT tested dies: 3', slide_text)
            self.assertIn('Class tested dies: 3', slide_text)
            self.assertIn('Either-test tested dies: 3', slide_text)
            for old_label in ('bottom half', 'top half', 'DISPO', 'FB1594', 'FB1591', 'MDK', 'scrap', 'release', 'Hold'):
                self.assertNotIn(old_label, slide_text + table_text)


if __name__ == '__main__':
    unittest.main()