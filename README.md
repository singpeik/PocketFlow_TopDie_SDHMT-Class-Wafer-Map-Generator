# Wafer Test Result Analyzer

This repository is a source-only distribution. Supply the approved local `config.yaml`, `query.sql`, Intel CA bundle, AQUA credentials, and IOE/GPU reference geometry privately before a live run. Manufacturing inputs, generated outputs, certificates, and machine-local configuration are intentionally excluded from publication. No prior repository history or datasets are included in this source snapshot.

Analyze SDHMT and Class results for selected IOE or GPU wafers, and locate interest bins using distinct map markers. Shared UI labels use product-neutral wafer terminology; IOE/GPU labels remain where they identify the selected geometry. New workbooks use `VID to Wafer&WLA` for the cross-reference tab. Existing Python entry points, local environment paths, and previously generated workbooks are retained for compatibility.

> IMPORTANT: Prepare the CATTS report using the download instructions below. The active workflow collects inputs as Step 1, prepares geometry as Step 2, filters CATTS as Step 3, fetches SDHMT as Step 4, fetches Class as Step 5, and reports as Step 6.

## Current Workflow

Run `python main.py` from the project directory with the approved environment active, or use `run_wafermap_generator.bat`. Explicitly select IOE or GPU, enter the wafer list, and choose the CATTS raw report. The app does not infer the product from wafer IDs. The batch launcher changes to its own directory and currently expects the environment at `C:\venvs\pocketflow-IOE-SDHMT-Class-wafermap-generator`; adjust that local setting or invoke your intended Python directly.

A single-wafer input analyzes that wafer only. A multi-wafer input analyzes only that run's selected wafers. Neither mode automatically merges earlier workbooks: consolidation is a separate explicit Python API operation, described below. IOE-named Python entry points remain compatible with both die types.

Output dates are added automatically. References below to `outputs/Final.xlsx` mean the logical base name; normal runs create `outputs/Final_YYYYMMDD.xlsx`. Use a unique `--final-workbook` and `--mrb-ppt-output` for each batch or retry to avoid reusing same-day defaults. Close Excel/PPT outputs before writeback and never update an active run's workbook concurrently.

```powershell
python main.py --final-workbook outputs/Final_GPU_batch1.xlsx --mrb-ppt-output outputs/GPU_batch1.pptx
```

## Legacy Pipeline Features

The following legacy lot-rollup/reject-summary helpers remain in the source but are not the workflow launched by `main.py`. Their product/time-filter GUI and artifact names are historical; use the current wafer workflow and reporting rules in this README instead.

-   **Interactive GUI**: A user-friendly Tkinter pop-up window to input:
    -   **Lot IDs**: Supports single or multiple IDs (comma or newline separated).
    -   **Products**: Filter by one or multiple product strings (e.g., `%8FST%`).
    -   **Time Window**: Select a lookback period in days or a specific date range.
-   **Dynamic SQL Generation**: Constructs complex database queries parsing `query.sql` using `OR` conditions for multi-lot tracking. Automatically extracts the top-level 4-character BOM `DEVREVSTEP` (e.g. `8PXJ`) identifiers for mapping metadata downstream.
-   **Automated Data Dimensional Modeling**:
    -   Computes deep pivots and lot groupings on Interface Bin (IB) and Functional Bin (FB) pass/fail datasets.
    -   Produces `SDHMT die level data.csv`, `wafer level rollup.xlsx`, and `lot level rollup.xlsx` populated deeply with explicit `WAFER_SCRIBE` and dynamic `DEVREVSTEP` tracking identifiers.
    -   Dynamically collapses sparse groupings and trims flat 0-count columns.
-   **Intelligent PowerPoint Report Generation**:
    -   Generates `SDHMT_Report.pptx` via `python-pptx` wrapped in a native 16:9 Widescreen context. Stamped natively with current generation date strings and active run filters.
    -   Outputs high-visibility pastel colorcoded IB & FB wafer heatmaps per lot and per wafer dynamically mapped against standard dictionaries (`all_product_maps.json`) driven seamlessly by the row's extracted `DEVREVSTEP` BOM identity. Features precisely controlled `MultipleLocator(1)` scaled axis tick marks logic. 
    -   Applies complex Table level validations: Auto-colored Red text counts for true fails. Safe test indicators (`IB1`, `FB9301`, `IB93`, etc.) are safely anchored in standard black. `IB6` & `FB6` (Skipped Test) geometries display in unique Tint Yellow cells and enforce black text constraints.
    -   Includes built-in interactive slide-to-slide OpenXML Hyperlinking from summary tables down to dimensional detail slides for simple drill-down navigation.
-   **Drill-down Failure Summary Maps**:
    -   Produces `SDHMT reject bin summary.pptx` (Fail Bin Summary PPT) mapping specific wafers returning critical Interface Bin failures natively in 16:9 widescreen format. Excludes healthy markers and non-true fails (purges all `IB6` items).
    -   Constrains explicit red alert boundary coloring to accurately overlay precisely upon exact failure matrix targets (`s=40`).
    -   Generates explicitly printed cartesian failure coordinate pairs (X, Y) logically laid out beneath map heatmaps.
- **Automated Missing Map Detection & Ingestion**:
    - Includes a built in node `WaferMapCheckNode` executed directly after database querying. It intersects the extracted `DEVREVSTEP` BOM 4-character codes (e.g., `8PA9`) from the dataset and scans `all_product_maps.json`.
    - If a wafer background geometry map is missing for that specific BOM code, the execution intercepts and pops up a Tkinter instruction menu, prompting the user with exact steps to query CATTS and download an example product map `.xlsx` payload.
    - Generates a UI File picker window so the user can directly upload this missing map payload.
    - Intelligently assembles and constructs an overlapping union map of X, Y coordinates simultaneously parsed across all internal wafers in the `.xlsx` document; resolving empty dies omitted due to earlier uncaptured operations.
    - Stores the newly aggregated base layout footprint permanently under the missing BOM code directly to `all_product_maps.json`. Processing logically resumes seamlessly.

## Installation

1.  **Clone the repository**:
    ```bash
    git clone https://github.com/singpeik/PocketFlow_TopDie_SDHMT-Class-Wafer-Map-Generator.git
    cd PocketFlow_TopDie_SDHMT-Class-Wafer-Map-Generator
    ```

2.  **Request `all_product_maps.json`**:
    *   **CRITICAL**: The required `all_product_maps.json` file containing wafer geometries is proprietary and **not included** in this repository.
    *   After cloning, you must explicitly ask the author (**Yang1.Xiang**) for the `all_product_maps.json` file.
    *   Place the file exactly at this path: `domain-knowledge/wafer-map/all_product_maps.json` before proceeding.

3.  **Install dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

Live runs also require private PyUber access, the appropriate network/XEUS/AQUA permissions, local `config.yaml` and `query.sql`, your AQUA credentials, the approved `utils/IntelChain.pem` CA bundle, and the selected IOE/GPU reference geometry. Certificates and manufacturing data are not supplied by the source-only repository. Do not disable AQUA TLS verification.

## Usage Instructions

### Required Run Order

The workflow now enforces strict gating between steps to avoid crash cascades when AQUA is unstable.

Run `python main.py` for the complete sequence. Step 5 follows successful Steps 1-4, and Step 6 follows successful Class processing. If Step 5 fails, reporting stops. If there are no true VIDs, the workflow explicitly skips AQUA and can continue with blank Class bins.

Optional stop controls:
```bash
python main.py --stop-after-step 4
python main.py --stop-after-step 5
python main.py --stop-after-step 6
```

Important behavior:
- `main.py` default is `--stop-after-step 6` (full run).
- Stop controls end the run after the named step; `--stop-after-step 4` does not continue to Class automatically. These are not resume-only commands: another normal invocation starts input collection again.
- After Step 1-4, a popup shows pass/fail and row-level details.
- After Step 5, a popup shows pass/fail with fetched VID/row counts or failure reason.
- Step 6 validates Class columns and populated results, except for the explicit no-true-VID skip case.

### 1. Run the Application
Execute the main script from your terminal:
```bash
python main.py
```

### Step 1 Input Rules For Wafer Workflow

The Step 1 popup now supports both single-wafer and multi-wafer input.

Accepted input methods:
- Single `Lot7_WID3`, for example `C442Y19_017`
- Single `Lot8_WID3`, for example `C525Y9K0_102`
- Multi-wafer paste from an Excel column, one wafer per line
- Multi-wafer paste separated by tab, space, comma, or semicolon

Normalization behavior:
- `Lot7_WID3` is used directly after 3-digit WID normalization.
- `Lot8_WID3` is automatically converted to `Lot7_WID3` by trimming the lot token to 7 characters.
    - Example: `C525Y9K0_102` becomes `C525Y9K_102`
- If multiple wafers are pasted, the workflow normalizes every value and filters CATTS rows against the full normalized wafer list.

### 2. Step 2 CATTS Download BKM (Critical)

To avoid downloading the wrong CATTS raw report, follow this exact two-step method:

1. In CATTS, go to `Extended report` -> `MRB PRT Data extract reports` -> `QueryUnitData_UsingWafer`.
2. Input the target IOE or GPU wafers as normalized `Lot7_WID3` values.
3. Download the die-index report for all selected wafers.
4. In CATTS, go to `Query` -> `lot to unit integration`.
5. Input all selected wafer die indexes from the die-index report.
6. Manually select/type `ACW` and include exactly these two operations: `852060` and `216365`.
7. Download the CATTS raw Excel report and upload that file into the app.

Important:
- If step 1 is skipped, users can miss die indexes and produce incomplete ULT/CATTS raw data.
- If ACW operations are not explicitly set to `852060` and `216365`, the wrong report content may be downloaded.

### Legacy GUI Reference

This subsection and the missing-map subsection below describe the legacy pipeline only, not the current IOE/GPU input window.

A pop-up window will appear asking for parameters:
-   **Lot IDs**: Enter a list of Lot IDs you want to query (one lot per line, supporting paste from Excel).
-   **Products**: Enter your product BOM identifier (e.g. `%8P%` or specific code like `8PXJ`).
-   **Time Window**: Select either a rolling lookback period (e.g. last 600 days) or click "Specific Range" to define MM-DD-YYYY dates.
-   **Click "Run Analysis"** to start the pipeline data fetch from the server.

### 4. Handle Any Missing Maps (New Feature!)
After querying the data, the tool scans the payload for the specific `DEVREVSTEP` BOM 4-digit product identifier (e.g., `8PXJ` or `8FCM`). 
If a product map does **not** already exist in the local library (`all_product_maps.json`):
1. A warning prompt will pop up alerting you that the geometry is missing.
2. Follow the prompt instructions: Go to `CATTS` -> `Extended report` -> `die index reports` -> input your lot and download the `example-map` excel sheet to your machine.
3. Close the instruction prompt and the Native File Browser will appear.
4. Select that downloaded `example-map file` (e.g. `.xlsx` or `.csv`). The application will then dynamically learn the shape of the new wafer map directly from the uploaded file and automatically add it permanently to the library for flawless future runs without asking you again!

### Current Generated Artifacts

- `outputs/Final_YYYYMMDD.xlsx`: workbook containing `CATTS-cleaned`, `VID`, `WLA lots`, `VID to Wafer&WLA`, `VID to WLA`, `WLA lots SDHMT`, `AQUA class_7820`, and the Step 6 `Result Summary` tab.
- `outputs/AQUA class_7820_YYYYMMDD.xlsx`: local AQUA evidence/cache workbook.
- `outputs/IOE MRB map summaryYYYYMMDD.pptx` or `outputs/GPU MRB map summaryYYYYMMDD.pptx`: normal wafer-result report. Custom output names are supported.
- Explicit consolidated runs write a separate workbook and PPT; their inputs remain unchanged. Actual unique wafer counts can differ from requested list sizes if CATTS data is absent.

### 4. Step 3 and Step 4 Workbook Enrichment

Current Step 3 and Step 4 behavior writes all downstream-ready data into `outputs/Final.xlsx`.

Step 3 now writes these workflow tabs:
- `CATTS-cleaned`
- `VID`
- `WLA lots`
- `VID to Wafer&WLA`
- `VID to WLA`

Step 3 also enriches `CATTS-cleaned` with these lookup columns:
- `WLA lot`
- `WLA wafer`
- `WLA die index`

Step 4 uses the Step 3 W-lot list to fetch XEUS SDHMT data and writes the result into `outputs/Final.xlsx` tab `WLA lots SDHMT`.

Important Step 4 matching rule:
- `CATTS-cleaned`.`WLA die index` is the lookup key.
- Step 4 writes `SDHMT IB` and `SDHMT FB` back into `CATTS-cleaned` by exact `WLA die index` match.
- The old `VISUAL_ID` to SDHMT lookup description is no longer correct for the wafer workflow.

### 5. AQUA Class Data Fetch

This repo includes a reusable AQUA utility for fetching Class `7820` data from the true VID list stored in `outputs/Final.xlsx` tab `VID` during Step 5 of the main workflow.

Important clone/setup note:
- Files under `domain-knowledge/wafer-map/` are local-only and are not pushed with the normal repo clone.
- A new user must reach out to the repo owner to get `domain-knowledge/wafer-map/ioe_wafer_map.json`.
- A new user must also reach out to the repo owner to get the BKM file `domain-knowledge/wafer-map/AQUA-fetch-guide.md`.
- A new user must create their own local `domain-knowledge/wafer-map/AQUA API Key.yaml` using their own AQUA user ID, API key, and environment.

What is required:
- AQUA user ID
- AQUA REST API key
- the local credentials file `domain-knowledge/wafer-map/AQUA API Key.yaml`
- `outputs/Final.xlsx` with true VID values stored in tab `VID`
- the privately supplied approved CA chain `utils/IntelChain.pem` for Python HTTPS verification

What was learned from the sample AQUA client:
- The GitHub reference `https://github.com/intel-sandbox/aqua-rest-api-client` was used to learn the host mapping and the request order: generate, poll, optionally inspect config, then execute.
- The sample client was only a reference. The final working utility in this repo still handles the environment-specific parts: proxy bypass, Intel CA bundle trust, UNC-style execute path, and recovery of the server-side ZIP when direct writeback fails.

Current reusable entry points:
- Utility module: `utils/aqua_client.py`
- Standalone CLI: `fetch_aqua_class_data.py`

Reason Step 5 remains explicitly gated:
- AQUA fetch depends on external network path stability.
- In unstable networks, Step 5 may return transient DNS/HTTP errors.
- The flow stops at Step 5 on failure and does not run Step 6.

Default workflow output:
- `outputs/Final.xlsx` with tab `AQUA class_7820`
- standalone cache workbook `outputs/AQUA class_7820.xlsx`

How it works now:
1. Read true VIDs from `outputs/Final.xlsx` tab `VID` in the workflow path.
2. Reuse an existing dated AQUA cache only when its `Visual_ID` set matches the current `VID` sheet and it includes `LATO_Valid_Flag` and `Operation_Final_Latest_Flag`. Old caches without the flags are bypassed. Matching cache contents do not prove live-system freshness after new testing; fetch fresh evidence when a retest is suspected.
3. Otherwise create the AQUA HTTPS session using `utils/IntelChain.pem` as the CA bundle.
4. Build the AQUA Class `7820` request in memory.
5. Generate and poll the AQUA job.
6. Execute using a UNC-style temporary path.
7. If direct writeback fails, recover the server-side ZIP path from the AQUA job status message.
8. Extract the temporary tab-delimited result.
9. Write the final AQUA Class result into `outputs/Final.xlsx` tab `AQUA class_7820` and also save the standalone cache workbook `outputs/AQUA class_7820.xlsx`.

The raw AQUA `.txt` file is treated as temporary working data and is not kept as the normal final deliverable.

Example standalone command:

```powershell
python fetch_aqua_class_data.py --workbook outputs/Final_YYYYMMDD.xlsx --output outputs/AQUA_class_review.xlsx
```

The standalone script fetches and saves AQUA evidence/cache; it does not itself merge Class bins into `CATTS-cleaned`. To refresh an existing final workbook programmatically, fetch fresh evidence, apply `apply_aqua_class_bins_to_final_workbook(...)`, then regenerate the Excel summary and PPT. Keep a backup. `--vid-csv` is an alternative input for a separate CSV-driven AQUA export.

Best-known method for new users cloning this repo:
- Reach out to the repo owner and request the local-only files under `domain-knowledge/wafer-map/`, especially `ioe_wafer_map.json` and `AQUA-fetch-guide.md`.
- Keep your AQUA credentials local in `domain-knowledge/wafer-map/AQUA API Key.yaml`.
- Create `domain-knowledge/wafer-map/AQUA API Key.yaml` yourself using your own AQUA `user_id`, `api_key`, and `environment` values.
- Use the approved privately supplied CA chain at `utils/IntelChain.pem`.
- Use `outputs/Final.xlsx` tab `VID` as the default true-VID input source.
- Only use `--vid-csv` when you intentionally want a separate CSV-driven fetch.

Current local documentation:
- Local-only BKM guide: `domain-knowledge/wafer-map/AQUA-fetch-guide.md` and must be requested from the repo owner after clone.
- Local credentials file: `domain-knowledge/wafer-map/AQUA API Key.yaml` and must be created by each user with their own AQUA account information.

Default final workbook artifact for the workflow:
- `outputs/Final.xlsx`

### 6. Wafer Test Result Reporting

Step 6 uses the final workbook's `CATTS-cleaned` sheet after SDHMT and AQUA Class writeback. It generates descriptive IOE/GPU results only, without disposition advice.

IOE map prerequisite:
- A fresh clone does not include `domain-knowledge/wafer-map/ioe_wafer_map.json`.
- A new user must request that local IOE wafer-map JSON from the repo owner before running the full IOE workflow.

Required Step 6 outputs:
The input window now includes an IOE/GPU geometry selector. Select one product per run before entering its wafer IDs. The selected map is prepared after input and before CATTS processing. GPU uses the same CATTS operations, SDHMT lookup, Class 7820 lookup, and interest-bin rules as IOE.

GPU reference geometry is generated from `domain-knowledge/wafer-map/GPU wafer map file.xlsx` into `domain-knowledge/wafer-map/gpu_wafer_map.json`. The current reference combines 12 wafers into 2,763 unique observed positions. Its outermost layer has 257 positions and its second layer has 253: the current two-layer edge band therefore has 510 positions, with 2,253 outside that band. The IOE reference has 2,577 positions. Counts depend on the supplied geometry and are not universal constants. Both products have separate entries in `all_product_maps.json`; building GPU does not replace IOE. X/Y-only Excel sources are supported. Existing JSON is reused; explicitly call `build_ioe_wafer_map(map_name="GPU", force_rebuild=True)` after replacing reference Excel to refresh the GPU JSON and library entry.

Optional GPU path overrides: `--gpu-map-source` and `--gpu-map-output`. GPU runs use GPU slide labels and a dated `GPU MRB map summary` PPT filename; IOE runs keep their IOE filename.

- The dated product-specific PPT and `Result Summary` tab in the dated final workbook.
- Whole-wafer interest-failure counts and two-layer exterior edge-band failure confirmation. No new MDK exports are generated.

Required PPT structure:
- The consolidated failure overview on slide 2 uses the exact detailed column order: `LOT7_WID3`, `Total Die Tested`, `Test %`, `SDHMT Tested Dies`, `SDHMT Whole Fails`, `SDHMT Edge Fails`, `SDHMT IB15 FB1590-1599`, `SDHMT IB36`, `Class Tested Dies`, `Class Whole Fails`, `Class Edge Fails`, `Class IB36`, `Combined Whole Fails`, `Combined Edge Fails`. It includes only wafers with positive combined whole failures and reads the detailed quantities from the full result summary.
- For a shorter consolidated PPT, pass `combined_maps_only=True` to keep only the SDHMT+Class map per tested wafer. Title, failure overview tables and detailed summaries remain; links are redirected to the retained combined maps. Existing source presentations are preserved when converting to a new output file.
- Reports with more than 15 wafer rows include the detailed failure-only table immediately after the title, sorted by combined whole failures descending. Its exact 14-column layout is listed above. Small normal runs use the regular summary page after the title instead of adding an extra failure overview. Detailed summary pages and complete Excel rows remain available.
- For consolidated reports, `consolidate_final_workbooks([...], new_output_path)` merges finalized workbooks without changing the inputs, removes stale source summary tabs, records source filenames, and keeps the later workbook's row for duplicate dies on the same wafer. `generate_ioe_mrb_summary_ppt(..., tested_wafers_only=True)` then keeps all wafers in Excel but omits zero-test wafers from the entire PPT, including detail maps. Tested wafers without interest failures remain included and summaries paginate rather than dropping them. Never consolidate a workbook while another run is still writing it.
- Summary slides hold up to 15 wafer rows. If the full list fits, keep every wafer. On overflow, first omit zero-Test-% rows; only if still too long, omit wafers with no interest-bin failure and display `Only wafers with interest-bin failures are shown in this table.` If the failing-wafer list still exceeds 15 rows, continue it across summary slides without dropping any failures. The Excel summary and all per-wafer detail slides always remain complete.
- The final workbook also contains a `Result Summary` tab with one row per wafer and the full result breakdown. It is an Excel table with filter/sort dropdowns, a frozen header and wafer-ID column, numeric `0.0%` formatting, and the same red/orange failure highlighting as the PPT. It includes the coverage columns, per-test and combined whole/edge counts, and all IB15/IB36 category counts. Reruns replace only this summary tab; other workbook sheets remain intact. Close the workbook in Excel before reporting so the atomic save can replace it safely.
- The summary table includes `Wafer Map Dies`, `Total Die Tested`, and `Test %` for each wafer. `Total Die Tested` counts unique dies with an eligible SDHMT or Class result once; `Test % = Total Die Tested / unique active positions in the selected wafer map`. Percentages use one decimal place. The GPU reference currently supplies a denominator of 2,763, not the historical fixed 2,400.
- First slide must be the summary/title slide.
- Title must be `IOE wafer test result summary` or `GPU wafer test result summary` for the selected product.
- Include current date and `PocketFlow Author: Yang1.Xiang/ SingPei`.
- The second slide summary table lists tested-die quantities, whole-wafer interest-failure counts, edge counts, and edge-failure Yes/No for SDHMT, Class, and their union. Each map slide also shows its tested-die quantity. Tested means a populated SDHMT IB other than IB6, or a populated Class IB; combined tested quantity counts unique dies from either test once. Duplicate die coordinates are counted once, with die index or visual ID used when coordinates are unavailable.
- For each selected IOE/GPU `LOT7_WID3`, normal reporting generates 3 slides; combined-only mode retains just the combined slide:
    1. SDHMT result slide
    2. Class result slide
    3. SDHMT + Class combined result slide

Required wafer-map slide formatting:
- Show X and Y coordinate axes on each wafer map.
- Use readable integer tick spacing that adapts to the current IOE wafer-map bounds.
- Use visible grey background die cells for the full IOE map geometry.
- Plot electrical results as markers at actual die X/Y coordinates instead of full-cell fills.
- IB15 with FB1590-1599 uses an upward triangle; IB15 with any other or missing FB uses a downward triangle; IB36 uses a diamond. SDHMT markers are red, Class markers are purple, good dies use a tint-green circle, and other bins remain background-only grey.
- The application visualizes observed test results only. It does not recommend scrap, release, or Hold for any unit or wafer.

Interest-bin rule (independently for SDHMT and Class):
- Include any `IB36` result, regardless of functional bin.
- Include every `IB15` result, regardless of functional bin. Each detail slide breaks down IB15 into FB1590-1599 inclusive, other FB values, and missing/unavailable FB values, alongside IB36 counts.
- Combined counts are the union of failing dies from either test, not the sum of test counts.
- Combined category counts are separate unique-die unions. A die with different categories in SDHMT and Class can appear in both category counts, so the category sum can exceed the combined unique interest-failure count.

Edge rule:
- Use the selected IOE/GPU reference geometry, not only the tested dies. Multi-wafer references use the union of observed positions; missing active positions are not invented.
- A die is in the outermost ring when an immediate horizontal or vertical neighbor belongs to exterior empty space. Empty positions are flood-filled from outside the reference bounds; enclosed internal holes do not create edge dies. This rule applies to both IOE and GPU geometry.
- Report whole-wafer and edge fail counts plus edge-failure Yes/No. If a failing die has missing or out-of-geometry coordinates, report `Unknown` instead of `No` unless another known edge failure already confirms `Yes`.
- There are no top/bottom-half metrics, disposition decisions, or new MDK files.
- The edge region is now two die layers wide: the exterior perimeter plus one adjacent horizontal/vertical active-die layer inward. Enclosed holes do not seed edge bands. The 257-die figure above describes only the original outermost layer.
- Summary tables and wafer captions emphasize positive whole-wafer interest-failure counts in bold red and positive edge-band counts/Yes indicators in bold orange. Zero counts, No/Unknown indicators, tested quantities, and Test % remain neutral; these colors are not scrap/release advice.

Combined-map precedence rule:
- AQUA Class bins are matched only from operation `7820` rows with both `LATO_Valid_Flag` and `Operation_Final_Latest_Flag` affirmative. LATO validity alone does not distinguish an earlier failure from a final pass. Latest/valid metadata is retained in the AQUA sheet; conflicting final-latest bins stop the merge rather than selecting the first row or always preferring pass. Old caches without the required flags are bypassed and refetched. Existing reports created with the former first-row rule must be refreshed before relying on their Class counts.
- Keep an interest failure from either test even if the other test is good or fails a different bin.
- When both tests fail the interest rule, show both markers at the same coordinate, with a larger red SDHMT marker behind the purple Class marker.
- For dies without an interest failure, keep the existing Class-first good-die rendering.

Focused regression checks: `python -m unittest discover -s tests -v`.

### Consolidation API Example

Consolidation and shorter PPTs are explicit Python API options, not `main.py` command-line switches. Combine only completed workbooks for the same product geometry. Put the intended authoritative input last: its entire die row wins on overlap. The utility copies tabular sheet values, records sources, and regenerates summaries; it does not guarantee preservation of source workbook formatting or formulas.

```python
from utils.ioe_workflow import consolidate_final_workbooks
from utils.reporting import generate_ioe_mrb_summary_ppt

combined = consolidate_final_workbooks(
    ["outputs/Final_GPU_batch1.xlsx", "outputs/Final_GPU_batch2.xlsx"],
    "outputs/Final_GPU_consolidated.xlsx",
)
report = generate_ioe_mrb_summary_ppt(
    final_workbook_path=combined["output_workbook"],
    ioe_map_path="domain-knowledge/wafer-map/gpu_wafer_map.json",
    output_ppt_path="outputs/GPU_consolidated_tested.pptx",
    output_dir="outputs",
    tested_wafers_only=True,
    combined_maps_only=True,
)
print(report["output_ppt"])
```

Both report options default to `False`. Tested-only mode omits zero-test wafers from the entire PPT but retains all Excel rows and all tested wafers, including tested wafers without interest failures. Summary pages paginate rather than removing those tested wafers. Combined-only mode preserves the overview/summaries and redirects links to combined maps. Existing PPTs can be converted in memory with `retain_combined_map_slides(presentation)` and saved under a new filename without fetching data again.

### Safe Reruns And Progress

- Watch `[Step X/6]` messages and `Plotting wafer N/total` in the terminal. The app does not provide a continuous overall percentage. Check for popups awaiting OK before assuming a process is stuck.
- Changing source during a run does not update that process's already-loaded code. New features need a new run or a deliberate backfill after completion.
- Do not interrupt Excel saves or update files belonging to an active run. Close Excel/PowerPoint outputs during writeback and use fresh filenames. A `BadZipFile`/CRC error requires preserving the damaged file and retrying into a fresh output, not repeatedly appending to it.
- A refreshed Excel Class result does not update previously generated PPT files. Regenerate the PPT before presenting new Class counts.
- Tests use synthetic inputs and cover latest-valid Class selection/cache rejection, unique tested counts, map-based percentages, two-layer exterior edges, exact slide-2 columns, tested-only filtering, combined-map links, consolidation, and Excel filtering/preservation/atomic-save recovery.
- Before committing, review the staged allowlist. Keep raw data, maps, results, credentials, certificates, config/SQL and local logs excluded. Ignore rules do not remove previously tracked private files or make unrelated Git history safe to publish.

Outputs folder retention:
- Keep `outputs/Final.xlsx` plus the dated final PPT. Previously generated MDK files are left untouched.
- Step 3 temporary CATTS extracts such as `filtered_catts_raw_*.xlsx` and `filtered_catts_summary_*.json` are removed automatically right after they are merged into `outputs/Final.xlsx`.
- `step6_temp` contents are removed automatically when Step 6 reporting completes.

## Project Structure

-   `main.py`: Current IOE/GPU wafer workflow CLI and run parameters. Use `python main.py --help` for supported options.
-   `generate_reject_summary.py`: Standalone script generating the secondary Drill-down PPT with exact failure coordinates mapping.
-   `nodes.py`: Current six-step nodes, validation gates and completion popups; legacy pipeline nodes are retained separately.
-   `flow.py`: Establishes the execution path topology.
-   `utils/reporting.py`: Engine for drawing Wafer Heatmaps (`matplotlib`), structuring OpenXML hyperlinks, formatting data colorations, and adapting rectangular 16:9 limits.
-   `utils/ioe_workflow.py`: Wafer normalization, reference maps, CATTS/SDHMT/Class writeback and workbook consolidation.
-   `utils/aqua_client.py`: AQUA fetching, latest/valid metadata and cache handling.
-   `tests/test_interest_bins.py`: Synthetic regression checks; no live AQUA/XEUS fetch is required.
-   `config.yaml` / `query.sql` / `all_product_maps.json`: Source extraction geometries and setup configurations.
