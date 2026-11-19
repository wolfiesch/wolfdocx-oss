"""Defaults and choice sets the CLI parser needs before any handler runs.

This module imports only the standard library, so building the argument parser
does not load the handler modules that own these values. Each owning module
re-exports its names from here.
"""

from __future__ import annotations

from pathlib import Path

ALLOWED_CATEGORIES = {
    "package",
    "parts",
    "text",
    "structure",
    "styles",
    "formatting",
    "features",
}

PACKAGED_VALIDATOR_PROJECT = (
    Path(__file__).resolve().parent / "_openxml_validator" / "WolfDocx.OpenXmlValidator.csproj"
)
SOURCE_VALIDATOR_PROJECT = (
    Path(__file__).resolve().parents[2]
    / "tools"
    / "WolfDocx.OpenXmlValidator"
    / "WolfDocx.OpenXmlValidator.csproj"
)
DEFAULT_VALIDATOR_PROJECT = (
    PACKAGED_VALIDATOR_PROJECT if PACKAGED_VALIDATOR_PROJECT.is_file() else SOURCE_VALIDATOR_PROJECT
)
FILE_FORMAT_VERSIONS = {
    "office2007",
    "office2010",
    "office2013",
    "office2016",
    "office2019",
    "office2021",
    "microsoft365",
}

DEFAULT_PROJECT_FILE = Path(".wolfdocx/project.json")

SCHEMA_ROOT = Path(__file__).resolve().parent / "schemas"
CONTRACT_SCHEMAS = {
    "contracted-core": SCHEMA_ROOT / "contracted-core-v1.schema.json",
    "gate-policy": SCHEMA_ROOT / "gate-policy-v1.schema.json",
    "project": SCHEMA_ROOT / "project-v1.schema.json",
    "release-suite": SCHEMA_ROOT / "release-suite-v1.schema.json",
}

DEFAULT_MANIFEST = Path("corpus/manifest.json")

RAW_XML_MAX_BYTES = 16 * 1024 * 1024
"""Default ceiling on the total stored bytes of the XML parts one raw XML report returns."""
