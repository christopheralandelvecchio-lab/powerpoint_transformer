"""
extract_pptx_relationships.py

Purpose:
    Backward-compatible import shim for relationship extraction users who want
    a top-level module name.

Scope:
    New implementation code lives under powerpoint_extractor.
    This shim should not contain domain logic.
"""

from powerpoint_extractor.csv_writer import write_relationships_csv
from powerpoint_extractor.extract_pptx_relationships import extract_connector_relationships
from powerpoint_extractor.models import ConnectorRelationship

__all__ = [
    "ConnectorRelationship",
    "extract_connector_relationships",
    "write_relationships_csv",
]
