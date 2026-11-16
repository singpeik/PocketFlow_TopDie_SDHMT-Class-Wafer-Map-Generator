from pocketflow import Flow
from nodes import FetchDataNode, SaveOutputNode, WaferMapCheckNode, DataProcessingNode, LotLevelRollupNode, ReportGenerationNode, GenerateRejectSummaryNode
from nodes import GenerateIOEWaferMapNode, CollectIOEInputsNode, FilterCATTSRawReportNode, FetchSDHMTDataNode, FetchClassDataNode, GenerateIOEMRBReportNode

def create_pyuber_flow():
    """Create and return the PyUber SQL execution flow."""
    # Create nodes
    fetch_data = FetchDataNode()
    save_results = SaveOutputNode()
    wafer_map_check = WaferMapCheckNode()
    process_data = DataProcessingNode()
    lot_rollup = LotLevelRollupNode()
    report_gen = ReportGenerationNode()
    reject_summary = GenerateRejectSummaryNode()

    # Connect nodes in sequence
    fetch_data >> save_results >> wafer_map_check >> process_data >> lot_rollup >> report_gen >> reject_summary
    # Create flow starting with fetch_data
    return Flow(start=fetch_data)


def create_ioe_flow(stop_after_step=None):
    """Create and return the shared IOE/GPU wafer test result workflow."""
    generate_ioe_map = GenerateIOEWaferMapNode()
    collect_inputs = CollectIOEInputsNode()
    filter_raw_report = FilterCATTSRawReportNode()
    fetch_sdhmt = FetchSDHMTDataNode()
    fetch_class = FetchClassDataNode()
    generate_mrb_report = GenerateIOEMRBReportNode()

    collect_inputs >> generate_ioe_map >> filter_raw_report >> fetch_sdhmt

    if stop_after_step == 4:
        return Flow(start=collect_inputs)

    fetch_sdhmt >> fetch_class

    if stop_after_step == 5:
        return Flow(start=collect_inputs)

    fetch_class - "default" >> generate_mrb_report
    return Flow(start=collect_inputs)
