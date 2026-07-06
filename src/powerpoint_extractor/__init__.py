"""
powerpoint_extractor package

Public package interface for PowerPoint object and relationship extraction.
"""

from .csv_writer import write_objects_csv, write_relationships_csv
from .extract_pptx_objects import extract_pptx_objects, extract_shapes
from .extract_pptx_relationships import extract_connector_relationships
from .models import ConnectorRelationship, SlideObject

__all__ = [
    "ConnectorRelationship",
    "SlideObject",
    "extract_connector_relationships",
    "extract_pptx_objects",
    "extract_shapes",
    "write_objects_csv",
    "write_relationships_csv",
]
