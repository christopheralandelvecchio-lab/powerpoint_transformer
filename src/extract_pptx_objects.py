"""
extract_pptx_objects.py

Purpose:
    Backward-compatible import shim for legacy code that imports from the
    previous top-level extract_pptx_objects module.

Scope:
    New implementation code lives under the powerpoint_extractor package.
    This shim should not contain domain logic.
"""

from powerpoint_extractor.csv_writer import write_objects_csv
from powerpoint_extractor.extract_pptx_objects import extract_pptx_objects, extract_shapes
from powerpoint_extractor.models import SlideObject

__all__ = [
    "SlideObject",
    "extract_pptx_objects",
    "extract_shapes",
    "write_objects_csv",
]
