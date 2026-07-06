# PowerPoint Object Extraction - Baseline

## A python utility for extracting native PowerPoint objects from a user provided pptx file and returns csv file of detected objects and relationships. 

```text
PowerPoint Transformer is a Python-based extraction tool that converts native PowerPoint (.pptx) diagrams into structured CSV datasets representing graphical objects and their relationships. The project is intended to support model extraction, graph analysis, and digital engineering workflows through a reproducible, test-driven architecture.
```

## End-User Command

```bash
python src/main.py <pptx_file> <output_dir>
```

`main.py` writes both outputs by default:

- `<input_file_stem>_objects.csv`
- `<input_file_stem>_relationships.csv`

## Module Layout

```text
src/
└── powerpoint_extractor
    ├── __init__.py
    ├── csv_writer.py
    ├── extract_pptx_objects.py
    ├── extract_pptx_relationships.py
    ├── main.py
    ├── models.py
    └── utils.py
```

## Component Responsibilities

- `extract_pptx_objects.py` owns native object extraction.
- `models.py` owns shared dataclasses and CSV schemas.
- `csv_writer.py` owns CSV serialization.
- `main.py` owns end-user orchestration.

## Baseline Test Coverage

A small batch of test cases have been currated for testing purposes. 

The manifest-driven test suite is included to run tests and requires expected results in the test_suite. 

[Test Suite](test_suite/)
[Test Results](results/)

## Offline Test Command

```bash
python3 tools/run_test_suite.py
```

## Documentation

[Software Design Description](docs/PowerPoint_Object_Extraction_SDD_v0.3.docx)
[Interface Control Document](docs/PowerPoint_Object_Extraction_ICD_v0.5.docx)
[Change Log](CHANGELOG.md)

## Current Status

Implemented

- Native PowerPoint object extraction
- Relationship extraction
- CSV serialization
- Automated regression test suite

In Progress

- Group detection
- Spatial text association
- Connector inference

## Planned Features

- TWC ingestion
- Metadata extraction
- Group inference
- Relationship graph generation
- Plugin architecture
- Release packaging