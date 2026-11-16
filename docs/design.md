# Design Doc: IOE SDHMT and Class Wafer Map Intake

> Please DON'T remove notes for AI

## Requirements

> Notes for AI: Keep it simple and clear.

1. User Story: As an IOE engineer, I want the app to guide me through the CATTS intake workflow, build a reusable IOE wafer map from a known wafer-map workbook, and filter the uploaded CATTS raw report to one or many selected IOE wafers so the downstream SDHMT and Class analysis can start from consistent inputs.
2. Input:
   - The fixed IOE wafer map workbook containing active die X/Y coordinates across 25 wafers.
   - A user-entered IOE wafer list that may contain `Lot7_WID3` or `Lot8_WID3` values.
   - A user-uploaded CATTS raw Excel report.
3. Processing:
   - Build a master IOE wafer map by unioning active X/Y die locations across the workbook.
   - Prompt the user for the target IOE wafer identifier list.
   - Accept single-wafer and multi-wafer paste input from Excel.
   - Accept values separated by newline, tab, space, comma, or semicolon.
   - Normalize both `Lot7_WID3` and `Lot8_WID3` user input into `Lot7_WID3` before filtering.
    - Enforce the Step 2 CATTS download BKM before file upload:
       - CATTS `Extended report` -> `MRB PRT Data extract reports` -> `QueryUnitData_UsingWafer` with selected `Lot7_WID3` values.
       - Download the IOE die-index report first.
       - CATTS `Query` -> `lot to unit integration` using those die indexes.
       - Manually set `ACW` operations to exactly `852060` and `216365`.
       - Download that CATTS raw Excel report and upload it into the app.
   - Prompt the user to upload the CATTS raw report after manual CATTS steps.
   - Filter the CATTS raw report to the selected wafer.
   - Enforce wafer ID normalization as a required parsing rule. I also fixed one real parsing bug in `ioe_workflow.py`: wafer IDs read from Excel as values like `17.0` are now normalized to `017` instead of `170`. That was the reason the first real test returned zero filtered rows.
   - Treat `FAB_WAFER`, `WAFER_ID`, and `SORT_WAFER_ID` as short 3-digit WID fields, so values like `17`, `17.0`, or `017` must all normalize to `017`.
   - Treat `WAFER_SCRIBE` as the long wafer ID field and do not normalize it to a 3-digit WID.
   - Write `outputs/Final.xlsx`, including the `VID`, `WLA lots`, `VID to IOE&WLA`, and `VID to WLA` tabs from the CATTS data.
   - Enrich `CATTS-cleaned` with `WLA lot`, `WLA wafer`, and `WLA die index` columns during Step 3.
   - Fetch XEUS SDHMT data for the W lot7% list, store the result in `outputs/Final.xlsx` tab `WLA lots SDHMT`, and match `CATTS-cleaned`.`WLA die index` to `WLA lots SDHMT`.`WLA die index`.
   - Fetch or reuse AQUA Class `7820` data from `outputs/Final.xlsx` tab `VID`, verify HTTPS using `utils/IntelChain.pem`, write the returned data into `outputs/Final.xlsx` tab `AQUA class_7820`, persist a standalone cache workbook `outputs/AQUA class_7820.xlsx`, write `Class IB` and `Class FB` back into `outputs/Final.xlsx`, and use the updated `CATTS-cleaned` data for IOE wafer-map reporting.
   - Generate a summary-first PowerPoint that includes per-wafer SDHMT, Class, and combined IOE wafer maps, fail metrics, test-rate calculations, disposition logic, and `scrap` MDK exports.
   - Preserve X/Y, VID, fake VID, IOE die index, and W lot references for later SDHMT and Class data fetches.
4. Output:
   - A generated IOE wafer map JSON file.
   - `outputs/Final.xlsx` with `CATTS-cleaned`, `VID`, `WLA lots`, `VID to IOE&WLA`, `VID to WLA`, `WLA lots SDHMT`, and `AQUA class_7820` tabs.
   - `outputs/AQUA class_7820.xlsx` as the reusable standalone Step 5 cache workbook.
   - `outputs/Final.xlsx` with `Class IB` and `Class FB` written back into `CATTS-cleaned`.
   - `IOE chipping MRB summary_YYYYMMDD.pptx` with summary and per-wafer wafer-map slides.
   - `{LOT7_WID3} MDK_YYYYMMDD.csv` for each wafer with `scrap` disposition.
   - A JSON summary listing filtered row counts and W lot references.

## Wafer ID SPEC

The content of this section is mandatory and must be enforced in the `.py` implementation.

1. **Short WID fields**
   - `FAB_WAFER`, `WAFER_ID`, and `SORT_WAFER_ID` are short wafer ID fields.
   - These fields must always normalize to a 3-digit WID.
   - Examples:
     - `17` -> `017`
     - `17.0` -> `017`
     - `017` -> `017`

2. **Long wafer ID field**
   - `WAFER_SCRIBE` is the long wafer ID field.
   - `WAFER_SCRIBE` must remain the full 12-character wafer identifier.
   - `WAFER_SCRIBE` must never be converted to a 3-digit WID.

3. **Mandatory code enforcement**
   - Any Python implementation that parses or exports these fields must enforce the short-WID normalization rule for `FAB_WAFER`, `WAFER_ID`, and `SORT_WAFER_ID`.
   - Any Python implementation that parses or exports `WAFER_SCRIBE` must preserve it as the long wafer ID.
   - Violating this section is a bug and must be treated as a spec regression.

## VID SPEC

The content of this section is mandatory and must be enforced in the `.py` implementation.

1. **Fake VID definition**
   - A fake VID is a `VISUAL_ID` value that contains `_`.
   - A fake VID is not a backend true VID.
   - A fake VID is the same identifier as the WLA die index.
   - Example: `AJPHI801MME2_2_-3`

2. **Fake VID usage**
   - Fake VID must be used only for SDHMT lookup.
   - Fake VID must be matched exactly against `WLA die index`.
   - If `CATTS-cleaned`.`WLA die index` equals `WLA lots SDHMT`.`WLA die index`, then:
     - `INTERFACE_BIN` must populate `SDHMT IB`
     - `FUNCTIONAL_BIN` must populate `SDHMT FB`

3. **True VID definition**
   - A true VID is a `VISUAL_ID` value that does not contain `_`.
   - A true VID is the backend unit identifier used for Class data fetch.

4. **True VID usage**
   - True VID must be extracted from the selected wafer rows and written to `Final.xlsx` tab `VID`.
   - `Final.xlsx` tab `VID` is the required input list for Class data fetch.
   - True VID must not be used as the lookup key for SDHMT `WLA die index` matching.

5. **Mandatory code enforcement**
   - Any Python implementation that classifies `VISUAL_ID` values must enforce the fake-VID versus true-VID split exactly as defined above.
   - Any Python implementation that populates `SDHMT IB` and `SDHMT FB` must use only fake VID exact matching to `WLA die index`.
   - Any Python implementation that prepares Class-fetch inputs must use only true VID values.
   - Violating this section is a bug and must be treated as a spec regression.

## Bin Hierarchy SPEC

The content of this section is mandatory and must be enforced in the `.py` implementation.

1. **IB definition**
   - `IB` means `INTERFACE_BIN`.
   - `IB` values must contain only 1 to 2 digits.
   - Examples: `IB1`, `IB8`, `IB15`, `IB19`, `IB97`.

2. **FB definition**
   - `FB` means `FUNCTIONAL_BIN`.
   - `FB` values must contain 3 to 4 digits.
   - Examples: `FB101`, `FB144`, `FB1591`, `FB9768`.

3. **IB to FB hierarchy rule**
   - `FB` is always a sub-category under an `IB` parent bin.
   - For a 3-digit `FB`, the parent `IB` is the first digit.
   - For a 4-digit `FB`, the parent `IB` is the first two digits.
   - Examples:
     - `FB101` belongs to `IB1`
     - `FB144` belongs to `IB1`
     - `FB1591` belongs to `IB15`
     - `FB9768` belongs to `IB97`

4. **Good die rule**
   - A good die is any die whose parent bin is `bin1`.
   - Therefore, a die is good if either:
     - `IB == 1`, or
     - `FB` belongs to parent `IB1`.
   - Examples of good dies include `IB1`, `FB101`, and `FB144`.

5. **Mandatory code enforcement**
   - Any Python implementation that determines whether a die is good must treat `IB1` and all `FB1xx` descendants as good dies.
   - Any Python implementation that groups or colors `FB` values must preserve the `FB` to `IB` parent-child relationship defined above.
   - Violating this section is a bug and must be treated as a spec regression.


## Flow Design

> Notes for AI:
> 1. Consider the design patterns of agent, map-reduce, rag, and workflow. Apply them if they fit.
> 2. Present a concise, high-level description of the workflow.

### Applicable Design Pattern:

Workflow: A linear intake pipeline that prepares the IOE geometry first, guides the user through the manual CATTS steps, filters the uploaded raw file into a clean wafer-specific dataset for one or more selected wafers, fetches SDHMT and AQUA Class results back into the final workbook, then generates IOE wafer-map reporting and disposition outputs.

### Flow high-level Design:

1. **Generate IOE Wafer Map Node**: Reads the 25-wafer IOE wafer-map workbook, unions active X/Y coordinates across wafers, and writes the reusable IOE base map into JSON.
2. **Collect IOE Inputs Node**: Pops up the Step 1 prompt for an IOE wafer list, accepts `Lot7_WID3` or `Lot8_WID3` input, supports Excel-column paste, normalizes the selection into `Lot7_WID3`, then shows the Step 2 CATTS BKM instruction window with the exact two-step path (`QueryUnitData_UsingWafer` first, then `lot to unit integration` with ACW operations `852060` and `216365`) before allowing CATTS raw Excel upload.
3. **Filter CATTS Raw Report Node**: Filters the uploaded CATTS raw report to the selected normalized `Lot7_WID3` list, writes `outputs/Final.xlsx`, stores the true VID list in the `VID` tab, builds `VID to IOE&WLA` and `VID to WLA`, enriches `CATTS-cleaned` with `WLA lot`, `WLA wafer`, and `WLA die index`, collects W lot7% values from `LOT_AT_216365` and `LOT_AT_852060`, and saves the Step 3 artifacts.
4. **Fetch SDHMT Data Node**: Executes the XEUS SDHMT query using the derived W lot7% list, creates `WLA die index`, appends the `WLA lots SDHMT` tab to `outputs/Final.xlsx`, and fills `SDHMT IB` and `SDHMT FB` back into `CATTS-cleaned` by exact `WLA die index` match.
5. **Fetch Class Data Node**: Uses `outputs/Final.xlsx` tab `VID` as the true-VID input list for AQUA API Class `7820` fetch. This step must use the local AQUA credentials YAML file, may reuse `outputs/AQUA class_7820.xlsx` only when its `Visual_ID` set exactly matches the current `VID` sheet, must write the returned data into `outputs/Final.xlsx` tab `AQUA class_7820`, must refresh the standalone cache workbook, must write `Class IB` and `Class FB` back into `outputs/Final.xlsx` `CATTS-cleaned` by exact `VISUAL_ID` match, and must not reuse the fake-VID SDHMT lookup logic.
6. **Generate IOE MRB Report Node**: Uses `outputs/Final.xlsx` `CATTS-cleaned` plus the IOE geometry to build a 5-slide reporting pattern: title page, summary table page, then per-wafer SDHMT, Class, and combined wafer-map slides with visible X/Y axes, grey background die cells, center-dot result rendering, plus `{LOT7_WID3} MDK_YYYYMMDD.csv` files for `scrap` wafers.

```mermaid
flowchart LR
   map[Generate IOE Wafer Map] --> input[Collect IOE Inputs]
   input --> filter[Filter CATTS Raw Report]
   filter --> sdhmt[Fetch SDHMT Data]
   sdhmt --> classFetch[Fetch Class Data]
   classFetch --> report[Generate IOE MRB Report]
```
## Utility Functions

> Notes for AI:
> 1. Understand the utility function definition thoroughly by reviewing the doc.
> 2. Include only the necessary utility functions, based on nodes in the flow.

1. **IOE Workflow Utilities** (`utils/ioe_workflow.py`)
   - *Input*: workbook paths, normalized `Lot7_WID3` list, uploaded CATTS raw report path
   - *Output*: IOE map JSON, filtered pandas DataFrame, summary JSON
   - *Necessity*: Centralizes the IOE-specific parsing, filtering, map generation, and GUI prompts.

2. **Tkinter Dialogs** (`utils/ioe_workflow.py`)
   - *Input*: prompt strings and optional initial directory
   - *Output*: normalized user selection values and uploaded file paths
   - *Necessity*: Implements the required Step 1 and Step 2 manual-user guidance workflow.

3. **XEUS SDHMT Fetch Helpers** (`utils/ioe_workflow.py`)
   - *Input*: W lot7% patterns, `config.yaml`, and `query.sql`
   - *Output*: SDHMT DataFrame and updated `outputs/Final.xlsx` with `WLA lots SDHMT`
   - *Necessity*: Implements Step 4 query execution and the `WLA die index` to `WLA die index` lookup.

4. **AQUA Client** (`utils/aqua_client.py`)
   - *Input*: `outputs/Final.xlsx` tab `VID`, `domain-knowledge/wafer-map/AQUA API Key.yaml`, `utils/IntelChain.pem`
   - *Output*: AQUA Class DataFrame for `outputs/Final.xlsx` tab `AQUA class_7820` and reusable cache workbook `outputs/AQUA class_7820.xlsx`
   - *Necessity*: Implements the Step 5 AQUA Class `7820` fetch flow, cache reuse check, Intel CA bundle trust, generate, poll, execute, server-ZIP recovery, and Excel conversion.

5. **Class Workbook Update Helper** (`utils/ioe_workflow.py`)
   - *Input*: `outputs/Final.xlsx`, AQUA Class DataFrame
   - *Output*: updated `CATTS-cleaned` sheet with `Class IB` and `Class FB`, plus `AQUA class_7820` tab
   - *Necessity*: Implements the Step 5 exact `VISUAL_ID` lookup from AQUA Class results back into the cleaned CATTS workbook.

6. **IOE MRB Reporting Utility** (`utils/reporting.py` or new IOE-specific reporting helper)
   - *Input*: `outputs/Final.xlsx`, IOE wafer geometry, and the user-selected normalized `LOT7_WID3` wafer list
   - *Output*: `IOE chipping MRB summary_YYYYMMDD.pptx` and `{LOT7_WID3} MDK_YYYYMMDD.csv`
   - *Necessity*: Implements the Step 6 title page, summary/disposition page, SDHMT/Class/combined wafer maps with visible coordinate axes, finalized color rules, fail counting rules, red/bold scrap-row emphasis on the summary page, and scrap-wafer MDK export.

## Node Design

### Shared Store

> Notes for AI: Try to minimize data redundancy

The shared store structure is organized as follows:

```python
shared = {
   "ioe_wafer_map": dict,             # Master IOE active die map
   "selected_wafer": dict,            # User-selected normalized Lot7_WID3 list and raw report path
   "filtered_catts_raw_df": pd.DataFrame,
   "filtered_catts_raw_summary": dict,
   "final_cleansed_df": pd.DataFrame,
   "wla_sdhmt_df": pd.DataFrame,
   "step4_summary": dict,
   "class_fetch_summary": dict,
   "step6_summary": dict
}
```

### Node Steps

> Notes for AI: Carefully decide whether to use Batch/Async Node/Flow.

1. GenerateIOEWaferMapNode
   - *Purpose*: Build the reusable IOE wafer map from the provided Excel workbook.
   - *Type*: Regular
   - *Steps*:
      - *prep*: Read the IOE wafer-map source path and target JSON paths from params.
      - *exec*: Load the workbook, union all valid active die X/Y coordinates across wafers, and write the IOE map JSON.
      - *post*: Store the generated map metadata in `shared`.

2. CollectIOEInputsNode
   - *Purpose*: Collect the target wafer selection and CATTS raw report path.
   - *Type*: Regular
   - *Steps*:
      - *prep*: Read the preferred initial directory for CATTS file selection.
      - *exec*: Show the Step 1 input prompt, accept one or many `Lot7_WID3` or `Lot8_WID3` values, normalize them into a `Lot7_WID3` list, then show the Step 2 instruction popup and collect the CATTS raw file upload path.
            - Step 2 popup must explicitly instruct users to first download IOE die indexes from `QueryUnitData_UsingWafer`, then use those die indexes in `lot to unit integration` with ACW operations `852060` and `216365`.
      - *post*: Store the normalized wafer list and raw report path in `shared`.

3. FilterCATTSRawReportNode
   - *Purpose*: Filter the uploaded CATTS raw report to the selected IOE wafer list and preserve the downstream identifiers.
   - *Type*: Regular
   - *Steps*:
      - *prep*: Read the normalized `Lot7_WID3` list, raw report path, and output directory.
      - *exec*: Normalize short wafer ID fields such as `FAB_WAFER` to 3-digit WID values, filter the raw Excel file against the selected `Lot7_WID3` list, classify fake versus true VIDs, collect W lot7% values, write `outputs/Final.xlsx`, and store the true VID list in the `VID` tab.
      - *post*: Store the filtered DataFrame and summary in `shared` and notify the user of the generated files.

4. FetchSDHMTDataNode
   - *Purpose*: Fetch SDHMT data for the W lot7% list and write the lookup results back into `outputs/Final.xlsx`.
   - *Type*: Regular
   - *Steps*:
      - *prep*: Read the Step 3 summary plus `config.yaml` and `query.sql` paths.
      - *exec*: Normalize short wafer ID fields such as `WAFER_ID` and `SORT_WAFER_ID` to 3-digit WID values, keep `WAFER_SCRIBE` as the long wafer ID, run the XEUS SDHMT query, create `WLA die index`, write the `WLA lots SDHMT` tab into `outputs/Final.xlsx`, and match only `CATTS-cleaned`.`WLA die index` values to populate `SDHMT IB` and `SDHMT FB`.
      - *post*: Store the enriched DataFrames and Step 4 summary in `shared` and notify the user of the fetch result.

   5. FetchClassDataNode
      - *Purpose*: Fetch Class `7820` data from AQUA using the true VID list from `outputs/Final.xlsx` tab `VID` and write Class bins back into `outputs/Final.xlsx`.
      - *Type*: Regular
      - *Steps*:
      - *prep*: Read `outputs/Final.xlsx`, the `VID` sheet name, the local AQUA credentials YAML path, and the Intel CA bundle path.
      - *exec*: Build the AQUA request from true VIDs only, reuse `outputs/AQUA class_7820.xlsx` only when its `Visual_ID` set exactly matches the current `VID` sheet, otherwise call AQUA generate and execute with HTTPS verification through `utils/IntelChain.pem`, recover the server ZIP if needed, convert the temporary AQUA result into a DataFrame, write that result into `outputs/Final.xlsx` tab `AQUA class_7820`, refresh the standalone cache workbook, then exact-match `VISUAL_ID` against `CATTS-cleaned` and populate `Class IB` and `Class FB` from `Unit_Interface_Bin` and `Unit_Functional_Bin`.
      - *post*: Store the AQUA Class-fetch result metadata in `shared`.

   6. GenerateIOEMRBReportNode
      - *Product geometry*: Choose IOE or GPU in the wafer-input window. Keep separate reference JSON files and named entries in the product-map library. GPU geometry accepts an X/Y-only Excel file; a wafer-ID column is optional. Use the selected geometry for edge detection, map backgrounds, and product labels. GPU uses the same CATTS, SDHMT, Class 7820, and interest-bin rules as IOE. The runtime executes input collection as Step 1 and geometry preparation as Step 2, before Step 3 CATTS processing; these supersede the original first-two-node order above.
      - *Purpose*: Generate Step 6 whole-wafer interest-bin reporting and confirm outermost edge-die failures for the selected wafers.
      - *Visualization objective*: Inspect SDHMT and Class test results for the input wafers, without scrap, release, or Hold recommendations. Plot IB15 with FB1590-1599 as triangles and IB36 as diamonds at their actual die coordinates. Use red for SDHMT and purple for Class. On combined maps retain both test markers on a shared die, with a larger SDHMT marker behind the Class marker so neither test is hidden. Keep legends consistent with the marker shapes and retain descriptive whole-wafer/edge counts.
      - *Shared naming*: Use Wafer Test Result Analyzer for the application and product-neutral wafer labels for common dialogs, input errors, and the VID to Wafer&WLA workbook tab. Preserve IOE/GPU names for product geometry, product-specific report labels, and compatible Python entry points and local environment paths.
      - *Tested die quantities*: Show unique SDHMT, Class, and either-test die counts per wafer in the summary table and detail slide captions. Preserve the existing tested definition: populated SDHMT IB other than IB6, or populated Class IB. Deduplicate using die X/Y, falling back to die index or visual ID when coordinates are unavailable. Count a die in both tests only once in the combined quantity. These counts describe available results, not wafer disposition.
      - *Expanded IB15 scope*: Include every IB15 result in SDHMT and Class, not only FB1590-1599. Distinguish IB15 FB1590-1599 with an upward triangle and other or missing FB with a downward triangle; retain IB36 diamonds. Detail slide captions break down IB15 FB1590-1599, other FB, missing FB, and IB36 counts. This rule supersedes the earlier restricted IB15 descriptions above. Combined counts are unique-die unions per category; a die can belong to different categories in different tests.
      - *Corrected outer perimeter*: Rebuild GPU geometry from the union of observed X/Y coordinates across all uploaded reference wafers. Do not treat enclosed missing positions as outside the wafer and do not invent active positions. Flood-fill empty grid positions from outside the geometry bounds using horizontal/vertical connectivity; only active dies adjacent to this exterior region are edge dies. This supersedes the earlier any-missing-neighbor rule. Force a JSON rebuild when replacing the reference spreadsheet, and synchronize the GPU entry in the product-map library without changing IOE geometry.
      - *Map-based test percentage*: Put Wafer Map Dies, Total Die Tested, and Test % next to the wafer ID in the PPT summary table. Total Die Tested is the unique either-test count; Test % is that quantity divided by the number of unique active positions in the selected reference geometry, formatted as a percentage. Do not use the historical fixed 2400-die denominator or double-count dies tested in both tests.
      - *Two-die edge width*: Use two die layers for IOE and GPU edge analysis: the exterior-connected perimeter and one adjacent horizontal/vertical layer inward. Expand only through observed active positions; do not seed edge bands from enclosed holes or add inferred die positions. Keep the one-layer option available for comparison. Update the PPT scope, edge counts, and geometry preview to identify this two-layer band.
      - *Failure font emphasis*: In summary tables and detail captions, render nonzero whole-wafer interest-failure and IB15/IB36 category counts in bold red. Render nonzero two-layer edge-band failure counts and confirmed edge-failure Yes indicators in bold orange. Keep zeros, No/Unknown indicators, tested quantities, and Test % neutral. These highlights describe observations only and are not disposition recommendations.
      - *Excel result summary*: Before plotting, write the same per-wafer result rows into a Result Summary tab in the final workbook, including all IB15/IB36 breakdowns. Put wafer ID, map die total, unique tested total and numeric percentage first. Use an Excel table with filters, freeze the header and wafer-ID column, format percentages as 0.0%, and apply the same failure font emphasis. Replace only this summary tab on reruns. Save to a sibling temporary workbook and atomically replace the final workbook only after a successful save, preserving other tabs and protecting the previous file on failure. Never modify a workbook belonging to an active concurrent run.
      - *Type*: Regular
      - *Steps*:
      - *prep*: Read `outputs/Final.xlsx`, the IOE wafer geometry, and the selected normalized `LOT7_WID3` values.
      - *exec*: Build a title page, a summary table, and SDHMT, Class, and combined maps per selected wafer. An interest failure is IB36, or IB15 with FB1590 through FB1599 inclusive, independently for either test. Edge dies are active coordinates missing at least one immediate horizontal or vertical neighbor in the selected product's full geometry, not in the tested subset. Count whole-wafer and edge failures per test and the union of both tests, counting each die once in the combined result. Show edge-failure Yes/No; do not split by wafer half, assign disposition, or generate MDK files. Use the marker shapes and overlap rendering defined above, with grey geometry cells and green good-die circles. Generate the existing dated PPT deliverable with a wafer test result summary title. Missing edge coordinates must not be treated as confirmed absence of edge failures.
      - *post*: Store the Step 6 reporting summary in `shared`.

