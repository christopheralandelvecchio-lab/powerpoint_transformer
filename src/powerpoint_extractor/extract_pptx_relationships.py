"""
extract_pptx_relationships.py

Purpose:
    Reusable relationship extraction library for deriving first-pass connector
    relationships from extracted PowerPoint object records.

Scope:
    Baseline v0.3 connector relationship extraction. This module consumes
    SlideObject records produced by extract_pptx_objects.py and produces
    ConnectorRelationship records for relationships.csv.

Primary Interface:
    extract_connector_relationships(objects: list[SlideObject])
        -> list[ConnectorRelationship]

Design Note:
    This module does not perform general PowerPoint object traversal. Object
    extraction remains the responsibility of extract_pptx_objects.py.
"""

from __future__ import annotations

import math
from typing import Optional

from .models import ConnectorRelationship, SlideObject


CONNECTOR_MAX_ENDPOINT_DISTANCE_IN = 0.35
LABEL_PROXIMITY_RADIUS_IN = 0.65
CONNECTOR_SHAPE_TYPES = {"LINE", "CONNECTOR"}
LAYOUT_LINE_NAME_TOKENS = {"DIVIDER", "SWIMLANE", "SEPARATOR", "BOUNDARY"}


def is_connector_object(obj: SlideObject) -> bool:
    """
    Determine whether a SlideObject should be treated as a connector candidate.

    Parameters
    ----------
    obj:
        Extracted PowerPoint object record.

    Returns
    -------
    bool
        True when the object appears to be a native line or connector.
    """

    shape_type = obj.shape_type.upper()
    name = obj.shape_name.upper()

    # Some PowerPoint drawings use native lines as visual layout aids rather
    # than semantic connectors. Baseline v0.4 excludes common structural line
    # names so swimlane dividers and page separators do not become
    # relationships.csv rows.
    if any(token in name for token in LAYOUT_LINE_NAME_TOKENS):
        return False

    return (
        any(token in shape_type for token in CONNECTOR_SHAPE_TYPES)
        or "CONNECTOR" in name
        or shape_type == "LINE"
    )


def has_geometry(obj: SlideObject) -> bool:
    """
    Check whether an object has the geometry needed for endpoint matching.
    """

    return None not in (obj.left_in, obj.top_in, obj.width_in, obj.height_in)


def bbox(obj: SlideObject) -> tuple[float, float, float, float]:
    """
    Return an object's normalized bounding box in inches.

    Returns
    -------
    tuple[float, float, float, float]
        Bounding box formatted as left, top, right, bottom.
    """

    left = float(obj.left_in or 0)
    top = float(obj.top_in or 0)
    right = left + float(obj.width_in or 0)
    bottom = top + float(obj.height_in or 0)

    return min(left, right), min(top, bottom), max(left, right), max(top, bottom)


def center(obj: SlideObject) -> tuple[float, float]:
    """
    Return the center point of an object's bounding box in inches.
    """

    left, top, right, bottom = bbox(obj)
    return (left + right) / 2, (top + bottom) / 2


def connector_endpoints(connector: SlideObject) -> tuple[tuple[float, float], tuple[float, float]]:
    """
    Estimate connector endpoints from extracted connector geometry.

    Notes
    -----
    python-pptx exposes connector position and size through the same geometry
    fields as other shapes. This baseline treats the upper-left corner as the
    start endpoint and the lower-right corner as the end endpoint. This is a
    first-pass approximation that can be refined later with lower-level OOXML
    connector endpoint metadata if needed.
    """

    left = float(connector.left_in or 0)
    top = float(connector.top_in or 0)
    width = float(connector.width_in or 0)
    height = float(connector.height_in or 0)

    return (left, top), (left + width, top + height)


def point_to_bbox_distance(point: tuple[float, float], candidate: SlideObject) -> float:
    """
    Calculate the shortest distance from a point to an object's bounding box.

    A point inside the bounding box has distance zero.
    """

    x, y = point
    left, top, right, bottom = bbox(candidate)

    dx = max(left - x, 0, x - right)
    dy = max(top - y, 0, y - bottom)

    return math.hypot(dx, dy)


def nearest_object(
    endpoint: tuple[float, float],
    candidates: list[SlideObject],
    excluded_ids: set[str],
) -> tuple[Optional[SlideObject], Optional[float]]:
    """
    Find the nearest non-connector object to a connector endpoint.

    Parameters
    ----------
    endpoint:
        Connector endpoint coordinate in inches.

    candidates:
        Objects on the same slide as the connector.

    excluded_ids:
        Object IDs that must not be considered for this endpoint.

    Returns
    -------
    tuple[Optional[SlideObject], Optional[float]]
        Nearest object and measured endpoint distance. Returns None values when
        no candidate object can be evaluated.
    """

    nearest: Optional[SlideObject] = None
    nearest_distance: Optional[float] = None

    for candidate in candidates:
        if candidate.object_id in excluded_ids:
            continue
        if is_connector_object(candidate) or not has_geometry(candidate):
            continue

        distance = point_to_bbox_distance(endpoint, candidate)

        if nearest_distance is None or distance < nearest_distance:
            nearest = candidate
            nearest_distance = distance

    return nearest, nearest_distance


def infer_relationship_type(connector: SlideObject) -> str:
    """
    Infer a simple relationship type from connector name and text.

    This is intentionally conservative. It creates a useful first-pass type
    without pretending to understand the full diagram semantics.
    """

    text = f"{connector.shape_name} {connector.text}".lower()

    if "control" in text or "auth" in text:
        return "control_or_auth_flow"
    if "dependency" in text:
        return "dependency"

    return "connector_flow"


def connector_label(connector: SlideObject, objects_on_slide: list[SlideObject]) -> str:
    """
    Find candidate label text for a connector relationship.

    Key algorithm points
    --------------------
    1. Use text directly attached to the connector when available.
    2. Otherwise, find nearby non-connector text-bearing objects.
    3. Return up to two nearest labels within the configured proximity radius.
    """

    if connector.text.strip():
        return connector.text.strip()

    c_x, c_y = center(connector)
    labels: list[tuple[float, str]] = []

    for obj in objects_on_slide:
        if obj.object_id == connector.object_id:
            continue
        if is_connector_object(obj) or not obj.text.strip() or not has_geometry(obj):
            continue

        o_x, o_y = center(obj)
        distance = math.hypot(o_x - c_x, o_y - c_y)

        if distance <= LABEL_PROXIMITY_RADIUS_IN:
            labels.append((distance, obj.text.strip()))

    labels.sort(key=lambda item: item[0])

    return " | ".join(label for _, label in labels[:2])


def extract_connector_relationships(objects: list[SlideObject]) -> list[ConnectorRelationship]:
    """
    Infer connector relationships from extracted slide objects.

    Parameters
    ----------
    objects:
        SlideObject records produced by extract_pptx_objects().

    Returns
    -------
    list[ConnectorRelationship]
        First-pass inferred connector relationship records.

    Key algorithm points
    --------------------
    1. Partition objects by slide.
    2. Identify connector candidates on each slide.
    3. Estimate connector start and end endpoints.
    4. Match each endpoint to the nearest non-connector object.
    5. Assign confidence based on endpoint distance.
    6. Preserve uncertainty in the notes field instead of silently dropping
       weak or incomplete relationships.
    """

    relationships: list[ConnectorRelationship] = []
    by_slide: dict[int, list[SlideObject]] = {}

    for obj in objects:
        by_slide.setdefault(obj.slide_number, []).append(obj)

    for slide_number in sorted(by_slide):
        objects_on_slide = by_slide[slide_number]
        connector_index = 0

        for connector in objects_on_slide:
            if not is_connector_object(connector) or not has_geometry(connector):
                continue

            connector_index += 1
            start, end = connector_endpoints(connector)

            source, source_distance = nearest_object(
                start,
                objects_on_slide,
                {connector.object_id},
            )
            target, target_distance = nearest_object(
                end,
                objects_on_slide,
                {connector.object_id, source.object_id if source else ""},
            )

            notes: list[str] = []

            if source is None or target is None:
                notes.append("endpoint_match_incomplete")
            if source_distance is not None and source_distance > CONNECTOR_MAX_ENDPOINT_DISTANCE_IN:
                notes.append("source_distance_exceeds_threshold")
            if target_distance is not None and target_distance > CONNECTOR_MAX_ENDPOINT_DISTANCE_IN:
                notes.append("target_distance_exceeds_threshold")

            distances = [d for d in (source_distance, target_distance) if d is not None]
            average_distance = (
                sum(distances) / len(distances)
                if distances
                else CONNECTOR_MAX_ENDPOINT_DISTANCE_IN
            )
            confidence = max(
                0.0,
                min(1.0, 1 - (average_distance / CONNECTOR_MAX_ENDPOINT_DISTANCE_IN)),
            )

            relationships.append(
                ConnectorRelationship(
                    relationship_id=f"S{slide_number:03d}_REL_{connector_index:04d}",
                    slide_number=slide_number,
                    connector_object_id=connector.object_id,
                    source_object_id=source.object_id if source else "",
                    target_object_id=target.object_id if target else "",
                    relationship_type=infer_relationship_type(connector),
                    source_endpoint=f"{round(start[0], 3)},{round(start[1], 3)}",
                    target_endpoint=f"{round(end[0], 3)},{round(end[1], 3)}",
                    confidence=round(confidence, 3),
                    label_text=connector_label(connector, objects_on_slide),
                    notes=";".join(notes),
                )
            )

    return relationships
