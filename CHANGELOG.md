# Change Log - Baseline Control Logic Test Expansion

## 20260709 Changes
**Modified**
- Replaced SDD with a version with diagrams
- Replaced ICD with a version with diagrams
- Changed documents to remove version from file name

**Added**
- Assets directory in docs which are the images for the documents

---

## 20260705 Changes

**Added**
- TC012_sensor_action_chain.pptx with expected object and relationship CSV files.
- TC013_and_gate_enable.pptx with expected object and relationship CSV files.
- TC014_or_gate_alarm.pptx with expected object and relationship CSV files.
- TC015_not_gate_inhibit.pptx with expected object and relationship CSV files.
- TC016_latch_feedback_control.pptx with expected object and relationship CSV files.
- TC017_fault_shutdown_logic.pptx with expected object and relationship CSV files.
- Test matrix v0.5 entries for control logic relationship scenarios.

**Modified**
- test_suite/TEST_SUITE_MANIFEST.csv updated to include TC012 through TC017.
- docs/PowerPoint_Object_Extraction_Test_Matrix_v0.5.xlsx generated from v0.4 with added control logic cases.

**Deprecated**
- None.

**Notes**
- Expected relationships represent connector-derived relationships the current extraction module is expected to detect.
- Semantic interpretation of logic gates remains outside this baseline and belongs to later inference modules.
