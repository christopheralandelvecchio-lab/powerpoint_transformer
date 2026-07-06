"""
models.py

Purpose:
    Shared domain models and CSV schema definitions for the PowerPoint
    extraction package.

Scope:
    This module contains data structures only. It does not read PowerPoint
    files, infer relationships, or write output files.

Design Note:
    Keeping shared models in their own module prevents object extraction and
    relationship extraction from depending on each other's implementation
    details. That separation is important as the project grows beyond Phase 1.
"""

from __future__ import annotations

from dataclasses import dataclass, asdict
from typing import Optional


EXPECTED_OBJECT_CSV_FIELDS = [
    "object_id",
    "slide_number",
    "parent_group_id",
    "shape_name",
    "shape_type",
    "text",
    "left_in",
    "top_in",
    "width_in",
    "height_in",
    "rotation",
]

EXPECTED_RELATIONSHIP_CSV_FIELDS = [
    "relationship_id",
    "slide_number",
    "connector_object_id",
    "source_object_id",
    "target_object_id",
    "relationship_type",
    "source_endpoint",
    "target_endpoint",
    "confidence",
    "label_text",
    "notes",
]


@dataclass
class SlideObject:
    """
    Structured record for one detected PowerPoint object.

    Parameters
    ----------
    object_id:
        Unique identifier assigned by the object extractor.

    slide_number:
        One-based slide number where the object was found.

    parent_group_id:
        ID of the group object that contains this object.
        None means the object is directly on the slide.

    shape_name:
        PowerPoint's internal name for the shape.

    shape_type:
        General PowerPoint shape category exposed by python-pptx.

    text:
        Text contained inside the shape, if any.

    left_in, top_in:
        Position of the object on the slide in inches.

    width_in, height_in:
        Size of the object in inches.

    rotation:
        Rotation angle of the object, if available.
    """

    object_id: str
    slide_number: int
    parent_group_id: Optional[str]
    shape_name: str
    shape_type: str
    text: str
    left_in: Optional[float]
    top_in: Optional[float]
    width_in: Optional[float]
    height_in: Optional[float]
    rotation: Optional[float]


@dataclass
class ConnectorRelationship:
    """
    Structured record for one inferred connector relationship.

    Parameters
    ----------
    relationship_id:
        Unique relationship identifier assigned by the relationship extractor.

    slide_number:
        One-based slide number where the relationship was inferred.

    connector_object_id:
        Object ID of the native PowerPoint connector or line that produced the
        relationship candidate.

    source_object_id, target_object_id:
        Object IDs inferred to be connected by the connector endpoints.
        Blank values mean an endpoint could not be matched.

    relationship_type:
        First-pass classification of the relationship.

    source_endpoint, target_endpoint:
        Connector endpoint coordinates in inches, formatted as "x,y".

    confidence:
        Simple distance-based confidence score from 0.0 to 1.0.

    label_text:
        Relationship label text found on the connector or near the connector.

    notes:
        Semicolon-separated extraction notes, warnings, or threshold flags.
    """

    relationship_id: str
    slide_number: int
    connector_object_id: str
    source_object_id: str
    target_object_id: str
    relationship_type: str
    source_endpoint: str
    target_endpoint: str
    confidence: float
    label_text: str
    notes: str


def dataclass_to_csv_row(record) -> dict:
    """
    Convert a dataclass instance into a CSV-friendly dictionary.

    Parameters
    ----------
    record:
        Dataclass instance to serialize.

    Returns
    -------
    dict
        Dictionary with None values converted to blank strings.
    """

    row = asdict(record)

    # CSV output should use blank cells for missing values instead of the
    # literal string "None" so downstream ingestion remains clean.
    for key, value in row.items():
        if value is None:
            row[key] = ""

    return row
