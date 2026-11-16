import pandas as pd
import json
from collections import deque
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
import os
from pathlib import Path
from tempfile import NamedTemporaryFile
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import traceback
from utils.ioe_workflow import add_date_stamp_to_path

COLORS = ['yellow', 'orange', 'plum', 'cyan', 'pink', 'lightblue', 'gold', 'salmon', 'tan', 'khaki', 'coral', 'lightgrey', 'thistle']
IOE_DIE_MARKER_SIZE = 100
IOE_DIE_DOT_MARKER_SIZE = 14
IOE_DIE_FAIL_DOT_MARKER_SIZE = 22
IOE_BACKGROUND_CELL_COLOR = '#D0D0D0'
IOE_BACKGROUND_EDGE_COLOR = '#A8A8A8'
CLASS_FAIL_COLOR = '#A855F7'
SDHMT_FAIL_COLOR = '#F4A6A6'
BIN1_DOT_COLOR = '#98FB98'

def generate_lot_level_rollup(wafer_path, lot_path):
    print(f'Reading {wafer_path} for lot_level_rollup...')
    ib_pivot = pd.read_excel(wafer_path, sheet_name='IB pivot')
    fb_pivot = pd.read_excel(wafer_path, sheet_name='FB pivot')
    
    # We want to keep DEVREVSTEP, so group by LOT and DEVREVSTEP
    ib_lot = ib_pivot.groupby(['LOT', 'DEVREVSTEP']).sum(numeric_only=True).reset_index()
    if 'WAFER_SCRIBE' in ib_lot.columns:
        ib_lot.drop(columns=['WAFER_SCRIBE'], inplace=True)
        
    fb_lot = fb_pivot.groupby(['LOT', 'DEVREVSTEP']).sum(numeric_only=True).reset_index()
    if 'WAFER_SCRIBE' in fb_lot.columns:
        fb_lot.drop(columns=['WAFER_SCRIBE'], inplace=True)
        
    combined_lot = pd.merge(ib_lot, fb_lot, on=['LOT', 'DEVREVSTEP'], how='outer').fillna(0)
    
    with pd.ExcelWriter(lot_path, engine='openpyxl') as writer:
        ib_lot.to_excel(writer, sheet_name='IB lot rollup', index=False)
        fb_lot.to_excel(writer, sheet_name='FB lot rollup', index=False)
        combined_lot.to_excel(writer, sheet_name='IB+FB lot rollup', index=False)

    print(f'Saved lot level rollup to {lot_path}')
    return combined_lot

def get_base_bin(val, is_fb):
    if pd.isnull(val): return None
    v_str = str(int(val))
    if is_fb:
        return v_str[:-2] if len(v_str) > 2 else v_str
    return v_str

def draw_wafer_map(fig_name, active_dies, die_data, title, val_col, color_map, is_fb=False, fb_data=None):
    fig, ax = plt.subplots(figsize=(6, 6))
    xs, ys = [], []
    for x, y in active_dies:
        xs.append(x)
        ys.append(y)
        
    ax.scatter(xs, ys, color='gainsboro', marker='s', s=400, edgecolor='white')
    
    for x, y in active_dies:
        val = die_data.get((x, y))
        if val is not None and pd.notnull(val):
            b = get_base_bin(val, is_fb)
            
            fb_val = fb_data.get((x, y)) if fb_data else die_data.get((x, y))
            is_blue_93 = False
            if fb_val is not None and pd.notnull(fb_val):
                fb_str = str(int(fb_val))
                if fb_str.startswith('93') and fb_str != '9301':
                    is_blue_93 = True

            if b == '1':
                c = '#98FB98' # Pale Green
            elif b == '6' or (is_fb and str(b).startswith('6')):
                c = '#FFFFE0' # Light Yellow / Tint Yellow
            elif is_blue_93 and (is_fb or b == '93'):
                c = '#ADD8E6' # Light Blue
            elif b == '93':
                c = '#E0E0E0' # Lighter Grey
            else:
                if b not in color_map:
                    color_map[b] = COLORS[len(color_map) % len(COLORS)]
                c = color_map[b]
                
            # Make functional bin names smaller so they don't overlap as badly
            fsize = 7 if is_fb else 11
            ax.text(x, y, str(int(val)), ha='center', va='center', fontsize=fsize, color='black', zorder=3, weight='bold')
            ax.scatter([x], [y], color=c, marker='s', s=550, edgecolor='white', zorder=2, alpha=0.65)
            
    ax.set_aspect('equal')
    ax.set_title(title, fontsize=12)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#888888')
    ax.spines['left'].set_color('#888888')
    ax.tick_params(axis='both', colors='#888888', labelsize=8)
    ax.xaxis.set_major_locator(MultipleLocator(1))
    ax.yaxis.set_major_locator(MultipleLocator(1))
    ax.set_xlabel('X Coordinate', fontsize=9, color='#555555', weight='bold')
    ax.set_ylabel('Y Coordinate', fontsize=9, color='#555555', weight='bold')
    plt.tight_layout()
    plt.savefig(fig_name, dpi=120)
    plt.close()

def _add_df_to_slide(slide, df, title=None, left=Inches(0.5), top=Inches(1.5), width=Inches(9.0), height=Inches(5.0), font_size=Pt(10), title_size=Pt(20), link_col=None, highlight_bincounts=False, row_style_fn=None):
    if title:
        # Instead of clobbering slide.shapes.title, we add a separate textbox
        # right above the table so that the main slide title remains intact.
        txBox = slide.shapes.add_textbox(left, top - Inches(0.4), width, Inches(0.4))
        p = txBox.text_frame.paragraphs[0]
        p.text = title
        p.font.size = title_size
        p.font.bold = True
        
    rows, cols = df.shape
    if rows == 0 or cols == 0: return {}
    lot_links = {}
    try:
        table_shape = slide.shapes.add_table(rows+1, cols, left, top, width, height)
        table = table_shape.table
        
        for c, col_name in enumerate(df.columns):
            if col_name == 'LOT': table.columns[c].width = Inches(1.5)
            elif col_name == 'Wafer Count': table.columns[c].width = Inches(0.8)
            elif col_name == 'WAFER_SCRIBE': table.columns[c].width = Inches(1.5)

            cell = table.cell(0, c)
            cell.text = str(col_name)
            for p in cell.text_frame.paragraphs:
                p.font.size = font_size
                
        for r in range(rows):
            for c in range(cols):
                val = df.iloc[r, c]
                cell = table.cell(r+1, c)
                col_name = str(df.columns[c])
                
                if isinstance(val, (list, tuple, set)):
                    val_text = ', '.join(str(x) for x in val)
                elif pd.isna(val):
                    val_text = ''
                elif isinstance(val, (int, float)):
                    if int(val) == val: val_text = str(int(val))
                    else: val_text = str(val)
                else:
                    val_text = str(val)
                    
                cell.text = ''
                p = cell.text_frame.paragraphs[0]
                run = p.add_run()
                run.text = val_text
                run.font.size = font_size

                if row_style_fn is not None:
                    style = row_style_fn(df.iloc[r], col_name, val_text) or {}
                    if style.get('bold'):
                        run.font.bold = True
                    if style.get('font_color'):
                        run.font.color.rgb = RGBColor.from_string(style['font_color'].replace('#', ''))
                
                if highlight_bincounts and val_text != '' and val_text != '0':
                    if col_name.startswith('IB') or col_name.startswith('FB'):
                        b = col_name[2:]
                        is_black = False
                        if col_name.startswith('IB'):
                            if b in ['1', '93', '6']: is_black = True
                        else:
                            if b == '1' or (len(b) == 3 and b.startswith('1')) or b == '9301' or str(b).startswith('6'): is_black = True
                        
                        if not is_black:
                            run.font.color.rgb = RGBColor(255, 0, 0)
                
                if link_col and col_name == link_col:
                    lot_links[val_text] = run

    except Exception as e:
        print(f'Error adding table to slide: {e}')
        
    return lot_links

def add_internal_link(run, slide_source, slide_target):
    run.hyperlink.address = 'dummy'
    ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    rId = slide_source.part.relate_to(slide_target.part, RT.SLIDE)
    run.hyperlink._hlinkClick.set(f'{{{ns}}}id', rId)
    run.hyperlink._hlinkClick.set('action', 'ppaction://hlinksldjump')

def generate_wafer_maps_and_ppt(die_csv, lot_xlsx, wafer_xlsx, map_json, output_ppt_path, filters=None):
    if filters is None:
        filters = {}
    print('Generating PPT Report...')
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
    except ImportError:
        print('python-pptx not installed.')
        return

    try:
        maps = json.load(open(map_json, encoding='utf-8'))
        fst8 = maps.get('8FST', {})
        active_dies = fst8.get('active_dies', [])
        
        die_df = pd.read_csv(die_csv)
        
        ib_wafer = pd.read_excel(wafer_xlsx, sheet_name='IB pivot')
        fb_wafer = pd.read_excel(wafer_xlsx, sheet_name='FB pivot')
        
        prs = Presentation()
        prs.slide_width = Inches(13.3333333)
        prs.slide_height = Inches(7.5)
        
        # Extract filters from passed-in dictionary
        import datetime
        try:
            products = filters.get('products', 'N/A')
            lots = filters.get('lots', 'N/A')
            time_cfg = filters.get('time', {})
            time_str = f"Mode: {time_cfg.get('mode', 'N/A')}, Days: {time_cfg.get('days', 'N/A')}"
        except:
            products, lots, time_str = 'N/A', 'N/A', 'N/A'
            
        today_str = datetime.date.today().strftime('%m/%d/%Y')
        
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = 'Pocketflow SDHMT wafer map generator'
        slide.placeholders[1].text = f'generated by pocketflow on {today_str}\npocketflow app author: Yang1.Xiang\n\nFilters Applied:\nProduct BOM: {products}\nLot: {lots}\nTime Range: {time_str}'
        for p in slide.placeholders[1].text_frame.paragraphs:
            p.font.size = Pt(14)
            
        slide_all = prs.slides.add_slide(prs.slide_layouts[6]) # Blank layout
        
        ib_lot_all = pd.read_excel(lot_xlsx, sheet_name='IB lot rollup')
        fb_lot_all = pd.read_excel(lot_xlsx, sheet_name='FB lot rollup')
        
        wafer_counts = ib_wafer.groupby('LOT')['WAFER_SCRIBE'].nunique().reset_index()
        wafer_counts.rename(columns={'WAFER_SCRIBE': 'Wafer Count'}, inplace=True)
        ib_lot_all = pd.merge(ib_lot_all, wafer_counts, on='LOT', how='left')
        
        # Ensure DEVREVSTEP is right after LOT if it exists
        cols = ['LOT']
        if 'DEVREVSTEP' in ib_lot_all.columns: cols.append('DEVREVSTEP')
        cols.append('Wafer Count')
        cols.extend([c for c in ib_lot_all.columns if c not in ['LOT', 'DEVREVSTEP', 'Wafer Count']])
        
        ib_lot_all = ib_lot_all[cols]
        
        for col in list(ib_lot_all.columns):
            if col not in ['LOT', 'DEVREVSTEP', 'Wafer Count']:
                if (ib_lot_all[col] == 0).all(): ib_lot_all.drop(columns=[col], inplace=True)
                
        # Do the same for FB lot all
        fb_lot_all = pd.merge(fb_lot_all, wafer_counts, on='LOT', how='left')
        fb_cols = ['LOT']
        if 'DEVREVSTEP' in fb_lot_all.columns: fb_cols.append('DEVREVSTEP')
        fb_cols.append('Wafer Count')
        fb_cols.extend([c for c in fb_lot_all.columns if c not in ['LOT', 'DEVREVSTEP', 'Wafer Count']])
        fb_lot_all = fb_lot_all[fb_cols]

        for col in list(fb_lot_all.columns):
            if col not in ['LOT', 'DEVREVSTEP', 'Wafer Count']:
                if (fb_lot_all[col] == 0).all(): fb_lot_all.drop(columns=[col], inplace=True)
                
        ib_links_run_map = _add_df_to_slide(slide_all, ib_lot_all, title='All Lots - IB & FB Summary', left=Inches(0.5), top=Inches(1.2), width=Inches(12.0), height=Inches(2.5), font_size=Pt(8), title_size=Pt(18), link_col='LOT', highlight_bincounts=True)
        fb_links_run_map = _add_df_to_slide(slide_all, fb_lot_all, title=None, left=Inches(0.5), top=Inches(4.5), width=Inches(12.0), height=Inches(2.5), font_size=Pt(8), link_col='LOT', highlight_bincounts=True)
        
        lots = die_df['LOT'].unique()
        lot_target_slides = {}

        for lot in lots:
            lot_ib_wafer = ib_wafer[ib_wafer['LOT'] == lot].copy()
            lot_fb_wafer = fb_wafer[fb_wafer['LOT'] == lot].copy()
            
            if 'LOT' in lot_ib_wafer.columns: lot_ib_wafer.drop(columns=['LOT'], inplace=True)
            if 'LOT' in lot_fb_wafer.columns: lot_fb_wafer.drop(columns=['LOT'], inplace=True)
            
            for col in list(lot_ib_wafer.columns):
                if col not in ['WAFER_SCRIBE', 'DEVREVSTEP']:
                    if (lot_ib_wafer[col] == 0).all(): lot_ib_wafer.drop(columns=[col], inplace=True)

            for col in list(lot_fb_wafer.columns):
                if col not in ['WAFER_SCRIBE', 'DEVREVSTEP']:
                    if (lot_fb_wafer[col] == 0).all(): lot_fb_wafer.drop(columns=[col], inplace=True)
            
            # Ensure WAFER_SCRIBE and DEVREVSTEP are at the front
            def reorder_lot_cols(df):
                cols = []
                if 'WAFER_SCRIBE' in df.columns: cols.append('WAFER_SCRIBE')
                if 'DEVREVSTEP' in df.columns: cols.append('DEVREVSTEP')
                cols.extend([c for c in df.columns if c not in ['WAFER_SCRIBE', 'DEVREVSTEP']])
                return df[cols]
                
            lot_ib_wafer = reorder_lot_cols(lot_ib_wafer)
            lot_fb_wafer = reorder_lot_cols(lot_fb_wafer)
            
            devrevstep_lot_val = 'Unknown'
            if 'DEVREVSTEP' in lot_ib_wafer.columns and not lot_ib_wafer.empty:
                devrevstep_lot_val = str(lot_ib_wafer['DEVREVSTEP'].iloc[0])
            
            slide_lot_summary = prs.slides.add_slide(prs.slide_layouts[5])
            lot_target_slides[lot] = slide_lot_summary
            
            try:
                title_shape = slide_lot_summary.shapes.title
                title_shape.text = f'Lot {lot} ({devrevstep_lot_val}) IB and FB summary'
                if title_shape.text_frame.paragraphs:
                    for p in title_shape.text_frame.paragraphs:
                        p.font.size = Pt(18)
            except Exception: pass
            
            _add_df_to_slide(slide_lot_summary, lot_ib_wafer, title='IB Summary', left=Inches(0.5), top=Inches(1.5), width=Inches(12.0), height=Inches(2.0), font_size=Pt(9), highlight_bincounts=True)
            _add_df_to_slide(slide_lot_summary, lot_fb_wafer, title='FB Summary', left=Inches(0.5), top=Inches(4.2), width=Inches(12.0), height=Inches(2.0), font_size=Pt(9), highlight_bincounts=True)

            wafers = die_df[die_df['LOT'] == lot]['WAFER_SCRIBE'].unique()
            for w_idx, wafer in enumerate(wafers):
                color_map = {}
                
                ib_stats = ib_wafer[(ib_wafer['LOT'] == lot) & (ib_wafer['WAFER_SCRIBE'] == wafer)].copy()
                fb_stats = fb_wafer[(fb_wafer['LOT'] == lot) & (fb_wafer['WAFER_SCRIBE'] == wafer)].copy()
                
                if 'LOT' in ib_stats.columns: ib_stats.drop(columns=['LOT'], inplace=True)
                if 'WAFER_SCRIBE' in ib_stats.columns: ib_stats.drop(columns=['WAFER_SCRIBE'], inplace=True)
                
                if 'LOT' in fb_stats.columns: fb_stats.drop(columns=['LOT'], inplace=True)
                if 'WAFER_SCRIBE' in fb_stats.columns: fb_stats.drop(columns=['WAFER_SCRIBE'], inplace=True)
                
                # We don't want DEVREVSTEP in the stats table underneath the map either
                if 'DEVREVSTEP' in ib_stats.columns: ib_stats.drop(columns=['DEVREVSTEP'], inplace=True)
                if 'DEVREVSTEP' in fb_stats.columns: fb_stats.drop(columns=['DEVREVSTEP'], inplace=True)
                
                for col in list(ib_stats.columns):
                    if (ib_stats[col] == 0).all(): ib_stats.drop(columns=[col], inplace=True)

                for col in list(fb_stats.columns):
                    if (fb_stats[col] == 0).all(): fb_stats.drop(columns=[col], inplace=True)
                
                w_df = die_df[(die_df['LOT'] == lot) & (die_df['WAFER_SCRIBE'] == wafer)]
                
                devrevstep_val = 'Unknown'
                if 'DEVREVSTEP' in w_df.columns and not w_df.empty:
                    devrevstep_val = str(w_df['DEVREVSTEP'].iloc[0])
                
                # Fetch specific active dies for this wafer's DEVREVSTEP (defaults to 8FST if not found)
                wafer_active_dies = maps.get(devrevstep_val, maps.get('8FST', {})).get('active_dies', active_dies)
                
                ib_map_dict = {}
                fb_map_dict = {}
                for _, row in w_df.iterrows():
                    x, y = row.get('SORT_X'), row.get('SORT_Y')
                    ib = row.get('INTERFACE_BIN')
                    fb = row.get('FUNCTIONAL_BIN')
                    if pd.notnull(x) and pd.notnull(y):
                        ib_map_dict[(int(x), int(y))] = ib
                        fb_map_dict[(int(x), int(y))] = fb
                        
                ib_fig = f'ib_map_{lot}_{wafer}.png'
                fb_fig = f'fb_map_{lot}_{wafer}.png'
                
                draw_wafer_map(ib_fig, wafer_active_dies, ib_map_dict, f'IB Map - Wafer {wafer} ({devrevstep_val})', 'IB', color_map, is_fb=False, fb_data=fb_map_dict)
                draw_wafer_map(fb_fig, wafer_active_dies, fb_map_dict, f'FB Map - Wafer {wafer} ({devrevstep_val})', 'FB', color_map, is_fb=True, fb_data=fb_map_dict)
                
                slide_maps = prs.slides.add_slide(prs.slide_layouts[5])
                
                try:
                    title_shape = slide_maps.shapes.title
                    title_shape.text = f'Lot {lot}, Wafer {wafer} | {devrevstep_val} Maps & Details'
                    if title_shape.text_frame.paragraphs:
                        for p in title_shape.text_frame.paragraphs:
                            p.font.size = Pt(18)
                except Exception: pass
                
                slide_maps.shapes.add_picture(ib_fig, Inches(0.5), Inches(1.3), width=Inches(5.0))
                slide_maps.shapes.add_picture(fb_fig, Inches(6.5), Inches(1.3), width=Inches(5.0))
                
                _add_df_to_slide(slide_maps, ib_stats, title=None, left=Inches(0.5), top=Inches(6.4), width=Inches(5.0), height=Inches(0.8), font_size=Pt(8), highlight_bincounts=True)
                _add_df_to_slide(slide_maps, fb_stats, title=None, left=Inches(6.5), top=Inches(6.4), width=Inches(5.0), height=Inches(0.8), font_size=Pt(8), highlight_bincounts=True)
                
                os.remove(ib_fig)
                os.remove(fb_fig)
                
        for lot, run in ib_links_run_map.items():
            target_slide = lot_target_slides.get(lot)
            if target_slide and run:
                try:
                    add_internal_link(run, slide_all, target_slide)
                    run.font.color.rgb = RGBColor(0, 0, 255)
                    run.font.underline = True
                except Exception as ex:
                    pass

        prs.save(output_ppt_path)
        print(f'Saved PPT to {output_ppt_path}')
        
    except Exception as e:
        print(f'Error drawing maps/PPT: {e}')
        traceback.print_exc()


def _normalize_lot7_wid3_from_row(row):
    fab_lot = str(row.get('FAB_LOT') or '').strip().upper()
    fab_wafer = str(row.get('FAB_WAFER') or '').strip().zfill(3)
    return f"{fab_lot[:7]}_{fab_wafer}"


def _coerce_bin_value(value):
    if pd.isna(value):
        return None
    try:
        return int(float(value))
    except Exception:
        return None


def _derive_parent_ib_from_fb(value):
    fb_value = _coerce_bin_value(value)
    if fb_value is None:
        return None
    fb_digits = str(fb_value)
    if len(fb_digits) == 3:
        return int(fb_digits[0])
    if len(fb_digits) == 4:
        return int(fb_digits[:2])
    return None


def _is_good_die(ib_value=None, fb_value=None):
    ib_bin = _coerce_bin_value(ib_value)
    if ib_bin == 1:
        return True
    parent_ib = _derive_parent_ib_from_fb(fb_value)
    return parent_ib == 1


def _is_sdhmt_tested(row):
    sdhmt_ib = _coerce_bin_value(row.get('SDHMT IB'))
    return sdhmt_ib is not None and sdhmt_ib != 6


def _is_class_tested(row):
    class_ib = _coerce_bin_value(row.get('Class IB'))
    return class_ib is not None


def _is_interest_fail(row, test_name):
    ib_value = _coerce_bin_value(row.get(f'{test_name} IB'))
    return ib_value in {15, 36}


def _is_combined_fail(row):
    return _is_interest_fail(row, 'SDHMT') or _is_interest_fail(row, 'Class')


def _interest_bin_category(row, test_name):
    ib_value = _coerce_bin_value(row.get(f'{test_name} IB'))
    if ib_value == 36:
        return 'IB36'
    if ib_value == 15:
        fb_value = _coerce_bin_value(row.get(f'{test_name} FB'))
        if fb_value is None:
            return 'IB15 Missing FB'
        return 'IB15 FB1590-1599' if 1590 <= fb_value <= 1599 else 'IB15 Other FB'
    return None


def _interest_bin_marker(row, test_name):
    return {'IB36': 'D', 'IB15 FB1590-1599': '^',
            'IB15 Other FB': 'v', 'IB15 Missing FB': 'v'}.get(_interest_bin_category(row, test_name))


def _find_edge_dies(active_dies, layers=2):
    if not isinstance(layers, int) or isinstance(layers, bool) or layers < 1:
        raise ValueError('Edge width must be a positive integer number of die layers.')
    geometry = set(active_dies)
    if not geometry:
        raise ValueError('The full wafer geometry is required to identify edge dies.')
    min_x = min(x for x, _ in geometry) - 1
    max_x = max(x for x, _ in geometry) + 1
    min_y = min(y for _, y in geometry) - 1
    max_y = max(y for _, y in geometry) + 1
    exterior = {(min_x, min_y)}
    pending = deque(exterior)
    while pending:
        x, y = pending.popleft()
        for neighbor in ((x - 1, y), (x + 1, y), (x, y - 1), (x, y + 1)):
            next_x, next_y = neighbor
            if (min_x <= next_x <= max_x and min_y <= next_y <= max_y
                    and neighbor not in geometry and neighbor not in exterior):
                exterior.add(neighbor)
                pending.append(neighbor)
    edge_dies = {
        (x, y) for x, y in geometry
        if any(neighbor in exterior for neighbor in (
            (x - 1, y), (x + 1, y), (x, y - 1), (x, y + 1)
        ))
    }
    frontier = edge_dies.copy()
    for _ in range(layers - 1):
        frontier = {
            neighbor for x, y in frontier
            for neighbor in ((x - 1, y), (x + 1, y), (x, y - 1), (x, y + 1))
            if neighbor in geometry and neighbor not in edge_dies
        }
        edge_dies.update(frontier)
        if not frontier:
            break
    return edge_dies


def _parse_dieindex_coordinates(intel_die_index):
    raw_value = str(intel_die_index or '').strip()
    if not raw_value:
        return (None, None)

    parts = raw_value.split('_')
    if len(parts) < 3:
        return (None, None)

    try:
        return int(parts[-2]), int(parts[-1])
    except ValueError:
        return (None, None)


def _resolve_plot_coordinates(row):
    x_value = row.get('X_LOCATION')
    y_value = row.get('Y_LOCATION')
    if pd.notna(x_value) and pd.notna(y_value):
        return int(x_value), int(y_value)

    dieindex_x, dieindex_y = _parse_dieindex_coordinates(row.get('INTEL_DIE_INDEX'))
    if dieindex_x is None or dieindex_y is None:
        return (None, None)
    return dieindex_x, dieindex_y


def _select_axis_tick_step(axis_span):
    if axis_span <= 0:
        return 1

    candidate_steps = [1, 2, 5, 10, 20, 25, 50]
    for step in candidate_steps:
        if axis_span / step <= 7:
            return step
    return candidate_steps[-1]


def _setup_ioe_map_axes(ax, active_dies, title):
    xs = [x for x, _ in active_dies]
    ys = [y for _, y in active_dies]
    ax.scatter(
        xs,
        ys,
        color=IOE_BACKGROUND_CELL_COLOR,
        marker='s',
        s=IOE_DIE_MARKER_SIZE,
        edgecolor=IOE_BACKGROUND_EDGE_COLOR,
        linewidths=0.35,
        alpha=0.98,
        zorder=1,
    )

    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)
    x_tick_step = _select_axis_tick_step(max_x - min_x)
    y_tick_step = _select_axis_tick_step(max_y - min_y)
    ax.set_xlim(min_x - 0.6, max_x + 0.6)
    ax.set_ylim(min_y - 0.6, max_y + 0.6)
    ax.set_aspect('equal', adjustable='box')
    ax.set_title(title, fontsize=12)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#888888')
    ax.spines['left'].set_color('#888888')
    ax.tick_params(axis='both', colors='#666666', labelsize=8)
    ax.xaxis.set_major_locator(MultipleLocator(x_tick_step))
    ax.yaxis.set_major_locator(MultipleLocator(y_tick_step))
    ax.set_xlabel('X Coordinate', fontsize=9, color='#555555', weight='bold')
    ax.set_ylabel('Y Coordinate', fontsize=9, color='#555555', weight='bold')


def _sdhmt_map_style(row):
    good_ib = _coerce_bin_value(row.get('SDHMT IB'))
    good_fb = _coerce_bin_value(row.get('SDHMT FB'))
    if _is_interest_fail(row, 'SDHMT'):
        return SDHMT_FAIL_COLOR
    if _is_good_die(good_ib, good_fb):
        return BIN1_DOT_COLOR
    return None


def _class_map_style(row):
    good_ib = _coerce_bin_value(row.get('Class IB'))
    good_fb = _coerce_bin_value(row.get('Class FB'))
    if _is_interest_fail(row, 'Class'):
        return CLASS_FAIL_COLOR
    if _is_good_die(good_ib, good_fb):
        return BIN1_DOT_COLOR
    return None


def _combined_style(row):
    class_fb = _coerce_bin_value(row.get('Class FB'))
    class_ib = _coerce_bin_value(row.get('Class IB'))
    sdhmt_fb = _coerce_bin_value(row.get('SDHMT FB'))
    sdhmt_ib = _coerce_bin_value(row.get('SDHMT IB'))

    if _is_interest_fail(row, 'Class'):
        return CLASS_FAIL_COLOR
    if _is_interest_fail(row, 'SDHMT'):
        return SDHMT_FAIL_COLOR

    if _is_class_tested(row):
        if _is_good_die(class_ib, class_fb):
            return BIN1_DOT_COLOR
        return None

    if _is_sdhmt_tested(row):
        if _is_good_die(sdhmt_ib, sdhmt_fb):
            return BIN1_DOT_COLOR
        return None

    return None


def _draw_ioe_map(fig_name, active_dies, wafer_df, title, style_resolver, test_name='SDHMT'):
    fig, ax = plt.subplots(figsize=(7.2, 7.2))
    _setup_ioe_map_axes(ax, active_dies, title)

    for _, row in wafer_df.iterrows():
        x, y = _resolve_plot_coordinates(row)
        fill_color = style_resolver(row)
        if x is None or y is None or fill_color is None:
            continue

        is_fail_color = fill_color in {CLASS_FAIL_COLOR, SDHMT_FAIL_COLOR}
        marker = _interest_bin_marker(row, test_name) or 'o'
        ax.scatter(
            [x],
            [y],
            color=fill_color,
            marker=marker,
            s=IOE_DIE_FAIL_DOT_MARKER_SIZE if is_fail_color else IOE_DIE_DOT_MARKER_SIZE,
            edgecolor='#555555' if is_fail_color else '#1F1F1F',
            linewidths=0.4 if is_fail_color else 0.3,
            zorder=3 if is_fail_color else 2,
            alpha=0.99,
        )

    plt.tight_layout(pad=0.15)
    plt.savefig(fig_name, dpi=180, bbox_inches='tight', pad_inches=0.05)
    plt.close()


def _draw_ioe_combined_map(fig_name, active_dies, wafer_df, title):
    fig, ax = plt.subplots(figsize=(7.2, 7.2))
    _setup_ioe_map_axes(ax, active_dies, title)

    for _, row in wafer_df.iterrows():
        x, y = _resolve_plot_coordinates(row)
        if x is None or y is None:
            continue

        sdhmt_marker = _interest_bin_marker(row, 'SDHMT')
        class_marker = _interest_bin_marker(row, 'Class')
        if sdhmt_marker or class_marker:
            for marker, color, size, layer in (
                (sdhmt_marker, SDHMT_FAIL_COLOR,
                 IOE_DIE_FAIL_DOT_MARKER_SIZE * (2.5 if class_marker else 1), 3),
                (class_marker, CLASS_FAIL_COLOR, IOE_DIE_FAIL_DOT_MARKER_SIZE, 4),
            ):
                if marker:
                    ax.scatter([x], [y], color=color, marker=marker, s=size,
                               edgecolor='#555555', linewidths=0.4, zorder=layer, alpha=0.99)
        else:
            fill_color = _combined_style(row)
            if fill_color:
                ax.scatter([x], [y], color=fill_color, marker='o', s=IOE_DIE_DOT_MARKER_SIZE,
                           edgecolor='#1F1F1F', linewidths=0.3, zorder=2, alpha=0.99)

    plt.tight_layout(pad=0.15)
    plt.savefig(fig_name, dpi=180, bbox_inches='tight', pad_inches=0.05)
    plt.close()


def _add_map_legend(slide, title, items, left=Inches(10.7), top=Inches(1.55), width=Inches(2.1)):
    title_box = slide.shapes.add_textbox(left, top, width, Inches(0.35))
    title_frame = title_box.text_frame
    title_frame.text = title
    for paragraph in title_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(14)
            run.font.bold = True

    current_top = top + Inches(0.5)
    for item in items:
        if len(item) == 3:
            label, hex_color, swatch_shape = item
        else:
            label, hex_color = item
            swatch_shape = MSO_SHAPE.RECTANGLE

        inverted_triangle = swatch_shape == 'v'
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ISOSCELES_TRIANGLE if inverted_triangle else swatch_shape,
            left, current_top, Inches(0.28), Inches(0.22),
        )
        if inverted_triangle:
            shape.rotation = 180
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(hex_color.replace('#', ''))
        shape.line.color.rgb = RGBColor(120, 120, 120)

        text_box = slide.shapes.add_textbox(left + Inches(0.38), current_top - Inches(0.02), width - Inches(0.38), Inches(0.28))
        text_frame = text_box.text_frame
        text_frame.text = label
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(13)
        current_top += Inches(0.4)


def _count_tested_dies(wafer_df, tested_mask):
    die_keys = set()
    for row_number, (_, row) in enumerate(wafer_df[tested_mask].iterrows()):
        x_value, y_value = _resolve_plot_coordinates(row)
        if x_value is not None and y_value is not None:
            die_keys.add(('xy', x_value, y_value))
            continue
        for column in ('INTEL_DIE_INDEX', 'VISUAL_ID'):
            value = row.get(column)
            if pd.notna(value) and str(value).strip():
                die_keys.add((column, str(value).strip()))
                break
        else:
            die_keys.add(('row', row_number))
    return len(die_keys)


def _build_step6_summary_row(lot7_wid3, wafer_df, active_dies):
    geometry = set(active_dies)
    edge_dies = _find_edge_dies(active_dies)
    coordinates = wafer_df.apply(_resolve_plot_coordinates, axis=1)
    sdhmt_mask = wafer_df.apply(lambda row: _is_interest_fail(row, 'SDHMT'), axis=1)
    class_mask = wafer_df.apply(lambda row: _is_interest_fail(row, 'Class'), axis=1)
    combined_mask = wafer_df.apply(_is_combined_fail, axis=1)
    summary = {'LOT7_WID3': lot7_wid3}
    sdhmt_tested = wafer_df.apply(_is_sdhmt_tested, axis=1)
    class_tested = wafer_df.apply(_is_class_tested, axis=1)
    tested_masks = {'SDHMT': sdhmt_tested, 'Class': class_tested,
                    'Combined': sdhmt_tested | class_tested}
    sdhmt_categories = wafer_df.apply(lambda row: _interest_bin_category(row, 'SDHMT'), axis=1)
    class_categories = wafer_df.apply(lambda row: _interest_bin_category(row, 'Class'), axis=1)
    for test_name, fail_mask in [('SDHMT', sdhmt_mask), ('Class', class_mask), ('Combined', combined_mask)]:
        summary[f'{test_name} Tested Dies'] = _count_tested_dies(wafer_df, tested_masks[test_name])
        fail_coordinates = set(coordinates[fail_mask])
        unknown_count = int((fail_mask & ~coordinates.map(lambda coord: coord in geometry)).sum())
        edge_count = len(fail_coordinates & edge_dies)
        summary[f'{test_name} Whole Fails'] = len(fail_coordinates & geometry) + unknown_count
        summary[f'{test_name} Edge Fails'] = edge_count
        summary[f'{test_name} Edge Fail?'] = 'Yes' if edge_count else ('Unknown' if unknown_count else 'No')
        for category in ('IB15 FB1590-1599', 'IB15 Other FB', 'IB15 Missing FB', 'IB36'):
            category_mask = ((sdhmt_categories == category) if test_name == 'SDHMT' else
                             (class_categories == category) if test_name == 'Class' else
                             ((sdhmt_categories == category) | (class_categories == category)))
            summary[f'{test_name} {category}'] = _count_tested_dies(wafer_df, category_mask)
    return {
        **summary,
        'Wafer Map Dies': len(geometry),
        'Total Die Tested': summary['Combined Tested Dies'],
        'Test %': summary['Combined Tested Dies'] / len(geometry),
    }


def _interest_failure_font_style(_row, column_name, value_text):
    if column_name.endswith(' Edge Fail?'):
        return {'bold': True, 'font_color': '#C65D00'} if value_text == 'Yes' else None
    count = _coerce_bin_value(value_text)
    if count is None or count <= 0:
        return None
    if column_name.endswith(' Edge Fails'):
        return {'bold': True, 'font_color': '#C65D00'}
    if (column_name.endswith(' Whole Fails') or ' IB15 ' in column_name
            or column_name.endswith(' IB36')):
        return {'bold': True, 'font_color': '#C00000'}
    return None


def write_wafer_result_summary(final_workbook_path, summary_rows):
    final_workbook = Path(final_workbook_path)
    summary_df = pd.DataFrame(summary_rows)
    coverage_columns = ['LOT7_WID3', 'Wafer Map Dies', 'Total Die Tested', 'Test %']
    if summary_df.empty:
        raise ValueError('No wafer results are available for the Excel summary.')
    summary_df = summary_df[coverage_columns + [column for column in summary_df.columns
                            if column not in coverage_columns and column != 'Combined Tested Dies']]
    workbook = load_workbook(final_workbook)
    temporary_path = None
    sheet_name = 'Result Summary'
    try:
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]
        sheet = workbook.create_sheet(sheet_name, 0)
        for values in dataframe_to_rows(summary_df, index=False, header=True):
            sheet.append(values)
        sheet.freeze_panes = 'B2'
        table_name = 'WaferResultSummary'
        existing_names = {name for worksheet in workbook for name in worksheet.tables}
        while table_name in existing_names:
            table_name += '_'
        table = Table(displayName=table_name, ref=sheet.dimensions)
        table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
        sheet.add_table(table)
        for column_index, column_name in enumerate(summary_df.columns, start=1):
            sheet.column_dimensions[get_column_letter(column_index)].width = min(max(len(column_name) + 2, 16), 30)
            sheet.cell(1, column_index).font = Font(bold=True, color='FFFFFF')
            for row_index in range(2, sheet.max_row + 1):
                cell = sheet.cell(row_index, column_index)
                if column_name == 'Test %':
                    cell.number_format = '0.0%'
                style = _interest_failure_font_style({}, column_name, str(cell.value))
                if style:
                    cell.font = Font(bold=True, color=style['font_color'].replace('#', ''))
        with NamedTemporaryFile(dir=final_workbook.parent, suffix='.xlsx', delete=False) as temporary:
            temporary_path = Path(temporary.name)
        workbook.save(temporary_path)
        workbook.close()
        os.replace(temporary_path, final_workbook)
    finally:
        workbook.close()
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
    return sheet_name


def _set_wafer_result_caption(text_frame, summary_row, test_name):
    prefix = 'Either-test' if test_name == 'Combined' else test_name
    lines = [
        [(f'{prefix} tested dies', 'Tested Dies'), ('interest fails', 'Whole Fails'),
         ('edge fails', 'Edge Fails'), ('edge fail?', 'Edge Fail?')],
        [('IB15 FB1590-1599', 'IB15 FB1590-1599'), ('IB15 other FB', 'IB15 Other FB'),
         ('IB15 missing FB', 'IB15 Missing FB'), ('IB36', 'IB36')],
    ]
    text_frame.clear()
    for line_index, metrics in enumerate(lines):
        paragraph = text_frame.paragraphs[0] if line_index == 0 else text_frame.add_paragraph()
        for metric_index, (label, suffix) in enumerate(metrics):
            key = f'{test_name} {suffix}'
            value = str(summary_row[key])
            run = paragraph.add_run()
            separator = ' ' if label.endswith('?') else ': '
            run.text = (' | ' if metric_index else '') + f'{label}{separator}{value}'
            run.font.size = Pt(14)
            style = _interest_failure_font_style(summary_row, key, value)
            if style:
                run.font.bold = style['bold']
                run.font.color.rgb = RGBColor.from_string(style['font_color'].replace('#', ''))


def generate_ioe_mrb_summary_ppt(final_workbook_path, ioe_map_path, output_ppt_path='IOE MRB map summary.pptx', output_dir='.'):
    final_workbook = Path(final_workbook_path)
    ioe_map = Path(ioe_map_path)
    output_ppt = add_date_stamp_to_path(output_ppt_path, separator='')
    output_ppt.parent.mkdir(parents=True, exist_ok=True)
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    cleaned_df = pd.read_excel(final_workbook, sheet_name='CATTS-cleaned')
    cleaned_df['LOT7_WID3'] = cleaned_df.apply(_normalize_lot7_wid3_from_row, axis=1)

    map_payload = json.loads(ioe_map.read_text(encoding='utf-8'))
    product_name = map_payload.get('map_name') or 'IOE'
    active_dies = [tuple(item) for item in map_payload.get('active_dies', [])]
    _find_edge_dies(active_dies)

    summary_rows = []

    prs = Presentation()
    prs.slide_width = Inches(13.3333333)
    prs.slide_height = Inches(7.5)

    wafer_groups = [(lot7_wid3, wafer_df.copy()) for lot7_wid3, wafer_df in cleaned_df.groupby('LOT7_WID3')]
    for _, wafer_df in wafer_groups:
        for col in ['SDHMT IB', 'SDHMT FB', 'Class IB', 'Class FB', 'X_LOCATION', 'Y_LOCATION']:
            if col in wafer_df.columns:
                wafer_df[col] = pd.to_numeric(wafer_df[col], errors='coerce')

    for lot7_wid3, wafer_df in wafer_groups:
        summary_row = _build_step6_summary_row(lot7_wid3, wafer_df, active_dies)
        summary_rows.append(summary_row)

    summary_sheet = write_wafer_result_summary(final_workbook, summary_rows)
    summary_df = pd.DataFrame(summary_rows)
    coverage_columns = ['LOT7_WID3', 'Wafer Map Dies', 'Total Die Tested', 'Test %']
    detail_columns = [column for column in summary_df.columns
                      if column not in coverage_columns and column != 'Combined Tested Dies'
                      and ' IB15 ' not in column and not column.endswith(' IB36')]
    summary_df = summary_df[coverage_columns + detail_columns]
    summary_df['Test %'] = summary_df['Test %'].map(lambda value: f'{value:.1%}')

    import datetime
    today_str = datetime.date.today().strftime('%m/%d/%Y')
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = f'{product_name} wafer test result summary'
    title_slide.placeholders[1].text = (
        f'Report generated by PocketFlow as of {today_str}\n'
        'PocketFlow Author: Yang1.Xiang/ SingPei\n\n'
        'Scope:\n'
        'Interest: all IB15 and IB36 results\n'
        'IB15: FB1590-1599 up-triangle; other/no FB down-triangle\n'
        'IB36: diamond\n'
        'Colors: SDHMT red; Class purple\n'
        'Edge: two die layers inward from exterior; enclosed gaps excluded\n'
        '1. Whole-wafer and edge-failure summary\n'
        '2. SDHMT wafer maps\n'
        '3. Class wafer maps\n'
        '4. Combined wafer maps'
    )
    for paragraph in title_slide.placeholders[1].text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(18)

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = f'{product_name} wafer test result summary'
    for paragraph in slide.shapes.title.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(22)

    summary_links = _add_df_to_slide(
        slide,
        summary_df,
        title=None,
        left=Inches(0.2),
        top=Inches(1.1),
        width=Inches(12.9),
        height=Inches(4.9),
        font_size=Pt(12),
        link_col='LOT7_WID3',
        row_style_fn=_interest_failure_font_style,
    )

    wafer_target_slides = {}

    for lot7_wid3, wafer_df in wafer_groups:
        summary_row = _build_step6_summary_row(lot7_wid3, wafer_df, active_dies)

        sdhmt_fb_fig = Path(output_dir) / f'{lot7_wid3}_sdhmt_fb.png'
        class_fb_fig = Path(output_dir) / f'{lot7_wid3}_class_fb.png'
        combo_fig = Path(output_dir) / f'{lot7_wid3}_combo.png'

        _draw_ioe_map(str(sdhmt_fb_fig), active_dies, wafer_df, f'{lot7_wid3} SDHMT Map', _sdhmt_map_style)
        _draw_ioe_map(str(class_fb_fig), active_dies, wafer_df, f'{lot7_wid3} Class Map', _class_map_style, test_name='Class')
        _draw_ioe_combined_map(str(combo_fig), active_dies, wafer_df, f'{lot7_wid3} Combined Map')

        slide1 = prs.slides.add_slide(prs.slide_layouts[5])
        wafer_target_slides[lot7_wid3] = slide1
        slide1.shapes.title.text = f'{product_name} {lot7_wid3} SDHMT result'
        for paragraph in slide1.shapes.title.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(18)
        slide1.shapes.add_picture(str(sdhmt_fb_fig), Inches(0.55), Inches(1.18), width=Inches(9.2), height=Inches(5.55))
        _add_map_legend(
            slide1,
            'Legend',
            [('IB15: FB1590-1599', SDHMT_FAIL_COLOR, MSO_SHAPE.ISOSCELES_TRIANGLE),
             ('IB15: other / no FB', SDHMT_FAIL_COLOR, 'v'),
             ('IB36', SDHMT_FAIL_COLOR, MSO_SHAPE.DIAMOND),
             ('Bin1 dot', BIN1_DOT_COLOR, MSO_SHAPE.OVAL), ('Other bins / untested cell', '#D3D3D3')],
            left=Inches(10.1),
            top=Inches(2.2),
            width=Inches(2.5),
        )
        tx1 = slide1.shapes.add_textbox(Inches(0.45), Inches(6.68), Inches(12.3), Inches(0.7))
        _set_wafer_result_caption(tx1.text_frame, summary_row, 'SDHMT')

        slide2 = prs.slides.add_slide(prs.slide_layouts[5])
        slide2.shapes.title.text = f'{product_name} {lot7_wid3} Class result'
        for paragraph in slide2.shapes.title.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(18)
        slide2.shapes.add_picture(str(class_fb_fig), Inches(0.55), Inches(1.18), width=Inches(9.2), height=Inches(5.55))
        _add_map_legend(
            slide2,
            'Legend',
            [('IB15: FB1590-1599', CLASS_FAIL_COLOR, MSO_SHAPE.ISOSCELES_TRIANGLE),
             ('IB15: other / no FB', CLASS_FAIL_COLOR, 'v'),
             ('IB36', CLASS_FAIL_COLOR, MSO_SHAPE.DIAMOND),
             ('Bin1 dot', BIN1_DOT_COLOR, MSO_SHAPE.OVAL), ('Other bins / untested cell', '#D3D3D3')],
            left=Inches(10.1),
            top=Inches(2.2),
            width=Inches(2.5),
        )
        tx2 = slide2.shapes.add_textbox(Inches(0.45), Inches(6.68), Inches(12.3), Inches(0.7))
        _set_wafer_result_caption(tx2.text_frame, summary_row, 'Class')

        slide3 = prs.slides.add_slide(prs.slide_layouts[5])
        slide3.shapes.title.text = f'{product_name} {lot7_wid3} SDHMT+Class result'
        for paragraph in slide3.shapes.title.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(18)
        slide3.shapes.add_picture(str(combo_fig), Inches(0.55), Inches(1.18), width=Inches(9.2), height=Inches(5.55))
        _add_map_legend(
            slide3,
            'Legend',
            [('SDHMT IB15 FB159x', SDHMT_FAIL_COLOR, MSO_SHAPE.ISOSCELES_TRIANGLE),
             ('SDHMT IB15 other/no FB', SDHMT_FAIL_COLOR, 'v'),
             ('SDHMT IB36', SDHMT_FAIL_COLOR, MSO_SHAPE.DIAMOND),
             ('Class IB15 FB159x', CLASS_FAIL_COLOR, MSO_SHAPE.ISOSCELES_TRIANGLE),
             ('Class IB15 other/no FB', CLASS_FAIL_COLOR, 'v'),
             ('Class IB36', CLASS_FAIL_COLOR, MSO_SHAPE.DIAMOND),
             ('Bin1 dot', BIN1_DOT_COLOR, MSO_SHAPE.OVAL), ('Other bins / untested cell', '#D3D3D3')],
            left=Inches(10.1),
            top=Inches(2.0),
            width=Inches(2.55),
        )
        tx3 = slide3.shapes.add_textbox(Inches(0.45), Inches(6.68), Inches(12.3), Inches(0.7))
        _set_wafer_result_caption(tx3.text_frame, summary_row, 'Combined')

        for fig_path in [sdhmt_fb_fig, class_fb_fig, combo_fig]:
            if fig_path.exists():
                fig_path.unlink()

    for lot7_wid3, run in summary_links.items():
        target_slide = wafer_target_slides.get(lot7_wid3)
        if target_slide is None:
            continue
        add_internal_link(run, slide, target_slide)
        run.font.underline = True
        run.font.color.rgb = RGBColor(0, 0, 255)

    prs.save(output_ppt)
    return {
        'output_ppt': str(output_ppt),
        'wafer_count': int(len(summary_rows)),
        'mdk_outputs': [],
        'summary_rows': summary_rows,
        'summary_sheet': summary_sheet,
        'final_workbook': str(final_workbook),
    }

