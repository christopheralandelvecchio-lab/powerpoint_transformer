"""
csv_writer.py

Purpose:
    CSV serialization functions for object and relationship extraction outputs.

Scope:
    This module writes already-extracted records to disk. It does not read
    PowerPoint files or infer relationships.
"""

from __future__ import annotations

import csv
from pathlib import Path

from .models import (
    ConnectorRelationship,
    EXPECTED_OBJECT_CSV_FIELDS,
    EXPECTED_RELATIONSHIP_CSV_FIELDS,
    SlideObject,
    dataclass_to_csv_row,
)


def write_objects_csv(objects: list[SlideObject], output_dir: Path, pptx_path: Path) -> Path:
    """
    Write extracted PowerPoint objects to an objects CSV file.

    Parameters
    ----------
    objects:
        List of extracted PowerPoint object records.

    output_dir:
        Directory where the CSV should be written.

    pptx_path:
        Original PowerPoint file path. Used to create the output file name.

    Returns
    -------
    Path
        Path to the generated objects CSV file.
    """

    output_dir = Path(output_dir)
    pptx_path = Path(pptx_path)

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{pptx_path.stem}_objects.csv"

    with output_path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=EXPECTED_OBJECT_CSV_FIELDS)
        writer.writeheader()

        for obj in objects:
            writer.writerow(dataclass_to_csv_row(obj))

    return output_path


def write_relationships_csv(
    relationships: list[ConnectorRelationship],
    output_dir: Path,
    pptx_path: Path,
) -> Path:
    """
    Write inferred connector relationships to a relationships CSV file.

    Parameters
    ----------
    relationships:
        List of inferred connector relationship records.

    output_dir:
        Directory where the CSV should be written.

    pptx_path:
        Original PowerPoint file path. Used to create the output file name.

    Returns
    -------
    Path
        Path to the generated relationships CSV file.
    """

    output_dir = Path(output_dir)
    pptx_path = Path(pptx_path)

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{pptx_path.stem}_relationships.csv"

    with output_path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=EXPECTED_RELATIONSHIP_CSV_FIELDS)
        writer.writeheader()

        for relationship in relationships:
            writer.writerow(dataclass_to_csv_row(relationship))

    return output_path
