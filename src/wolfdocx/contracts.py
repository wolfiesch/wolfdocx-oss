from __future__ import annotations

import os
import re
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Literal, Self, TypeVar

from . import _defaults
from .canonical import (
    is_portable_display_path,
    is_sha256,
    parse_utc_timestamp,
    semantic_sha256,
    serialize_utc_timestamp,
    sha256_bytes,
    strict_json_loads,
)

CONTRACT_SCHEMAS = _defaults.CONTRACT_SCHEMAS
SCHEMA_ROOT = _defaults.SCHEMA_ROOT

CONTRACTED_CORE_SCHEMA = CONTRACT_SCHEMAS["contracted-core"]
PROJECT_SCHEMA = CONTRACT_SCHEMAS["project"]
RELEASE_SUITE_SCHEMA = CONTRACT_SCHEMAS["release-suite"]


CONTRACTED_CORE_CONTRACT = "wolfdocx.contracted-core"
CONTRACTED_CORE_SCHEMA_VERSION = 1

_IDENTIFIER_PATTERN = re.compile(r"^[a-z][a-z0-9]*(?:[._-][a-z0-9]+)*$")
_SAFE_VALUE_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._+-]*$")
_MEDIA_TYPE_PATTERN = re.compile(r"^[a-z][a-z0-9!#$&^_.+-]*/[a-z0-9!#$&^_.+-]+$")


class ContractValidationError(ValueError):
    """Raised when a contracted-core value or wire object is invalid."""


def _invalid(field: str) -> ContractValidationError:
    return ContractValidationError(f"invalid {field}")


def _require_identifier(value: object, field: str) -> str:
    if (
        type(value) is not str
        or not value
        or len(value) > 128
        or _IDENTIFIER_PATTERN.fullmatch(value) is None
    ):
        raise _invalid(field)
    return value


def _require_safe_value(value: object, field: str) -> str:
    if (
        type(value) is not str
        or not value
        or len(value) > 128
        or _SAFE_VALUE_PATTERN.fullmatch(value) is None
    ):
        raise _invalid(field)
    return value


def _require_portable_display_path(value: object) -> str:
    if type(value) is not str or not is_portable_display_path(value):
        raise _invalid("display_path")
    return value


def _require_mapping(value: object) -> Mapping[str, object]:
    if not isinstance(value, Mapping) or any(type(key) is not str for key in value):
        raise ContractValidationError("invalid contract wire object")
    return value


def _require_exact_keys(value: Mapping[str, object], expected: frozenset[str]) -> None:
    if set(value) != expected:
        raise ContractValidationError("contract wire object has unsupported or missing fields")


def _require_list(value: object, field: str) -> list[object]:
    if type(value) is not list:
        raise _invalid(field)
    return value


_ItemT = TypeVar("_ItemT")

ToolOutcome = Literal["succeeded", "failed", "unavailable"]
DiagnosticSeverity = Literal["info", "warning", "error"]
ResultStatus = Literal["passed", "failed", "error"]


def _require_string(value: object, field: str) -> str:
    if type(value) is not str:
        raise _invalid(field)
    return value


def _require_int(value: object, field: str) -> int:
    if type(value) is not int:
        raise _invalid(field)
    return value


def _require_optional_string(value: object, field: str) -> str | None:
    if value is not None and type(value) is not str:
        raise _invalid(field)
    return value


def _require_optional_int(value: object, field: str) -> int | None:
    if value is not None and type(value) is not int:
        raise _invalid(field)
    return value


def _require_tool_outcome(value: object) -> ToolOutcome:
    if value == "succeeded":
        return "succeeded"
    if value == "failed":
        return "failed"
    if value == "unavailable":
        return "unavailable"
    raise _invalid("tool outcome")


def _require_diagnostic_severity(value: object) -> DiagnosticSeverity:
    if value == "info":
        return "info"
    if value == "warning":
        return "warning"
    if value == "error":
        return "error"
    raise _invalid("diagnostic severity")


def _require_result_status(value: object) -> ResultStatus:
    if value == "passed":
        return "passed"
    if value == "failed":
        return "failed"
    if value == "error":
        return "error"
    raise _invalid("result status")


def _require_typed_tuple(
    value: object,
    field: str,
    item_type: type[_ItemT],
) -> tuple[_ItemT, ...]:
    if type(value) is not tuple or any(not isinstance(item, item_type) for item in value):
        raise _invalid(field)
    return value


def _require_unique_roles(artifacts: tuple[ContractArtifact, ...], field: str) -> None:
    roles: set[str] = set()
    for artifact in artifacts:
        if artifact.role in roles:
            raise _invalid(field)
        roles.add(artifact.role)


def _deserialize_recorded_at(value: object) -> datetime:
    try:
        return parse_utc_timestamp(value)
    except ValueError as exc:
        raise _invalid("recorded_at") from exc


@dataclass(frozen=True, slots=True)
class ContractIdentity:
    contract: str
    schema_version: int
    operation: str
    operation_id: str

    def __post_init__(self) -> None:
        if type(self.contract) is not str or self.contract != CONTRACTED_CORE_CONTRACT:
            raise _invalid("contract")
        if (
            type(self.schema_version) is not int
            or self.schema_version != CONTRACTED_CORE_SCHEMA_VERSION
        ):
            raise _invalid("schema_version")
        _require_identifier(self.operation, "operation")
        _require_identifier(self.operation_id, "operation_id")

    def to_dict(self) -> dict[str, object]:
        return {
            "contract": self.contract,
            "schema_version": self.schema_version,
            "operation": self.operation,
            "operation_id": self.operation_id,
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, object]) -> Self:
        payload = _require_mapping(value)
        _require_exact_keys(
            payload,
            frozenset({"contract", "schema_version", "operation", "operation_id"}),
        )
        return cls(
            contract=_require_string(payload["contract"], "contract"),
            schema_version=_require_int(payload["schema_version"], "schema_version"),
            operation=_require_string(payload["operation"], "operation"),
            operation_id=_require_string(payload["operation_id"], "operation_id"),
        )


@dataclass(frozen=True, slots=True)
class ArtifactIdentity:
    sha256: str
    byte_count: int

    def __post_init__(self) -> None:
        if not is_sha256(self.sha256):
            raise _invalid("sha256")
        if type(self.byte_count) is not int or self.byte_count < 0:
            raise _invalid("byte_count")

    @classmethod
    def from_bytes(cls, value: bytes) -> Self:
        if not isinstance(value, bytes):
            raise _invalid("artifact bytes")
        return cls(sha256=sha256_bytes(value), byte_count=len(value))

    def to_dict(self) -> dict[str, object]:
        return {"sha256": self.sha256, "byte_count": self.byte_count}

    @classmethod
    def from_dict(cls, value: Mapping[str, object]) -> Self:
        payload = _require_mapping(value)
        _require_exact_keys(payload, frozenset({"sha256", "byte_count"}))
        return cls(
            sha256=_require_string(payload["sha256"], "sha256"),
            byte_count=_require_int(payload["byte_count"], "byte_count"),
        )


@dataclass(frozen=True, slots=True)
class ContractArtifact:
    role: str
    media_type: str
    display_path: str
    identity: ArtifactIdentity

    def __post_init__(self) -> None:
        _require_identifier(self.role, "role")
        if (
            type(self.media_type) is not str
            or len(self.media_type) > 255
            or _MEDIA_TYPE_PATTERN.fullmatch(self.media_type) is None
        ):
            raise _invalid("media_type")
        _require_portable_display_path(self.display_path)
        if not isinstance(self.identity, ArtifactIdentity):
            raise _invalid("artifact identity")

    def to_dict(self) -> dict[str, object]:
        return {
            "role": self.role,
            "media_type": self.media_type,
            "display_path": self.display_path,
            "identity": self.identity.to_dict(),
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, object]) -> Self:
        payload = _require_mapping(value)
        _require_exact_keys(
            payload,
            frozenset({"role", "media_type", "display_path", "identity"}),
        )
        return cls(
            role=_require_string(payload["role"], "role"),
            media_type=_require_string(payload["media_type"], "media_type"),
            display_path=_require_string(payload["display_path"], "display_path"),
            identity=ArtifactIdentity.from_dict(_require_mapping(payload["identity"])),
        )


@dataclass(frozen=True, slots=True)
class ToolIdentity:
    name: str
    version: str | None

    def __post_init__(self) -> None:
        _require_identifier(self.name, "tool name")
        if self.version is not None:
            _require_safe_value(self.version, "tool version")

    def to_dict(self) -> dict[str, object]:
        return {"name": self.name, "version": self.version}

    @classmethod
    def from_dict(cls, value: Mapping[str, object]) -> Self:
        payload = _require_mapping(value)
        _require_exact_keys(payload, frozenset({"name", "version"}))
        return cls(
            name=_require_string(payload["name"], "tool name"),
            version=_require_optional_string(payload["version"], "tool version"),
        )


@dataclass(frozen=True, slots=True)
class ContractTool:
    identity: ToolIdentity
    outcome: ToolOutcome
    exit_code: int | None

    def __post_init__(self) -> None:
        if not isinstance(self.identity, ToolIdentity):
            raise _invalid("tool identity")
        if type(self.outcome) is not str or self.outcome not in {
            "succeeded",
            "failed",
            "unavailable",
        }:
            raise _invalid("tool outcome")
        if self.exit_code is not None and type(self.exit_code) is not int:
            raise _invalid("tool exit_code")
        if self.outcome == "succeeded" and self.exit_code != 0:
            raise _invalid("tool exit_code")
        if self.outcome == "failed" and (type(self.exit_code) is not int or self.exit_code == 0):
            raise _invalid("tool exit_code")
        if self.outcome == "unavailable" and self.exit_code is not None:
            raise _invalid("tool exit_code")

    def to_dict(self) -> dict[str, object]:
        return {
            "identity": self.identity.to_dict(),
            "outcome": self.outcome,
            "exit_code": self.exit_code,
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, object]) -> Self:
        payload = _require_mapping(value)
        _require_exact_keys(payload, frozenset({"identity", "outcome", "exit_code"}))
        return cls(
            identity=ToolIdentity.from_dict(_require_mapping(payload["identity"])),
            outcome=_require_tool_outcome(payload["outcome"]),
            exit_code=_require_optional_int(payload["exit_code"], "tool exit_code"),
        )


@dataclass(frozen=True, slots=True)
class ContractDiagnostic:
    code: str
    severity: DiagnosticSeverity
    subject_role: str | None = None

    def __post_init__(self) -> None:
        _require_identifier(self.code, "diagnostic code")
        if type(self.severity) is not str or self.severity not in {
            "info",
            "warning",
            "error",
        }:
            raise _invalid("diagnostic severity")
        if self.subject_role is not None:
            _require_identifier(self.subject_role, "diagnostic subject_role")

    def to_dict(self) -> dict[str, object]:
        return {
            "code": self.code,
            "severity": self.severity,
            "subject_role": self.subject_role,
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, object]) -> Self:
        payload = _require_mapping(value)
        _require_exact_keys(payload, frozenset({"code", "severity", "subject_role"}))
        return cls(
            code=_require_string(payload["code"], "diagnostic code"),
            severity=_require_diagnostic_severity(payload["severity"]),
            subject_role=_require_optional_string(
                payload["subject_role"], "diagnostic subject_role"
            ),
        )


@dataclass(frozen=True, slots=True)
class ContractRequest:
    identity: ContractIdentity
    input_artifacts: tuple[ContractArtifact, ...] = ()

    def __post_init__(self) -> None:
        if not isinstance(self.identity, ContractIdentity):
            raise _invalid("request identity")
        artifacts = _require_typed_tuple(self.input_artifacts, "input_artifacts", ContractArtifact)
        _require_unique_roles(artifacts, "input_artifacts")

    def to_dict(self) -> dict[str, object]:
        return {
            "identity": self.identity.to_dict(),
            "input_artifacts": [artifact.to_dict() for artifact in self.input_artifacts],
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, object]) -> Self:
        payload = _require_mapping(value)
        _require_exact_keys(payload, frozenset({"identity", "input_artifacts"}))
        return cls(
            identity=ContractIdentity.from_dict(_require_mapping(payload["identity"])),
            input_artifacts=tuple(
                ContractArtifact.from_dict(_require_mapping(item))
                for item in _require_list(payload["input_artifacts"], "input_artifacts")
            ),
        )


@dataclass(frozen=True, slots=True)
class ContractResult:
    request: ContractRequest
    status: ResultStatus
    recorded_at: datetime
    output_artifacts: tuple[ContractArtifact, ...] = ()
    tools: tuple[ContractTool, ...] = ()
    diagnostics: tuple[ContractDiagnostic, ...] = ()

    def __post_init__(self) -> None:
        if not isinstance(self.request, ContractRequest):
            raise _invalid("result request")
        if type(self.status) is not str or self.status not in {
            "passed",
            "failed",
            "error",
        }:
            raise _invalid("result status")

        outputs = _require_typed_tuple(self.output_artifacts, "output_artifacts", ContractArtifact)
        _require_unique_roles(outputs, "output_artifacts")
        tools = _require_typed_tuple(self.tools, "tools", ContractTool)
        diagnostics = _require_typed_tuple(self.diagnostics, "diagnostics", ContractDiagnostic)

        if not isinstance(self.recorded_at, datetime) or self.recorded_at.tzinfo is None:
            raise _invalid("recorded_at")
        if self.recorded_at.utcoffset() != timedelta(0):
            raise _invalid("recorded_at")

        has_error_diagnostic = any(diagnostic.severity == "error" for diagnostic in diagnostics)
        has_unsuccessful_tool = any(tool.outcome != "succeeded" for tool in tools)
        if self.status == "passed" and (has_error_diagnostic or has_unsuccessful_tool):
            raise _invalid("passed result")
        if self.status == "failed" and (not has_error_diagnostic or has_unsuccessful_tool):
            raise _invalid("failed result")
        if self.status == "error" and not has_error_diagnostic:
            raise _invalid("error result")

    @property
    def identity(self) -> ContractIdentity:
        return self.request.identity

    def semantic_projection(self) -> dict[str, object]:
        """Return the explicit semantic projection that excludes only ``recorded_at``."""

        return {
            "request": self.request.to_dict(),
            "status": self.status,
            "output_artifacts": [artifact.to_dict() for artifact in self.output_artifacts],
            "tools": [tool.to_dict() for tool in self.tools],
            "diagnostics": [diagnostic.to_dict() for diagnostic in self.diagnostics],
        }

    @property
    def semantic_identity_sha256(self) -> str:
        """Return the identity of this result's explicit semantic projection."""

        return semantic_sha256(self.semantic_projection())

    def to_dict(self) -> dict[str, object]:
        return {
            "request": self.request.to_dict(),
            "status": self.status,
            "output_artifacts": [artifact.to_dict() for artifact in self.output_artifacts],
            "tools": [tool.to_dict() for tool in self.tools],
            "diagnostics": [diagnostic.to_dict() for diagnostic in self.diagnostics],
            "recorded_at": serialize_utc_timestamp(self.recorded_at),
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, object]) -> Self:
        payload = _require_mapping(value)
        _require_exact_keys(
            payload,
            frozenset(
                {
                    "request",
                    "status",
                    "output_artifacts",
                    "tools",
                    "diagnostics",
                    "recorded_at",
                }
            ),
        )
        return cls(
            request=ContractRequest.from_dict(_require_mapping(payload["request"])),
            status=_require_result_status(payload["status"]),
            output_artifacts=tuple(
                ContractArtifact.from_dict(_require_mapping(item))
                for item in _require_list(payload["output_artifacts"], "output_artifacts")
            ),
            tools=tuple(
                ContractTool.from_dict(_require_mapping(item))
                for item in _require_list(payload["tools"], "tools")
            ),
            diagnostics=tuple(
                ContractDiagnostic.from_dict(_require_mapping(item))
                for item in _require_list(payload["diagnostics"], "diagnostics")
            ),
            recorded_at=_deserialize_recorded_at(payload["recorded_at"]),
        )


def export_contract_schema(kind: str, output: str | Path) -> dict[str, object]:
    output_path = Path(output)
    source = CONTRACT_SCHEMAS[kind]
    if output_path.exists():
        return {
            "schema_version": 1,
            "passed": False,
            "contract": kind,
            "output": str(output_path),
            "sha256": None,
            "issues": ["schema output already exists"],
        }
    try:
        encoded = source.read_bytes()
        strict_json_loads(encoded)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        temporary = output_path.with_name(f".{output_path.name}.tmp")
        temporary.write_bytes(encoded)
        os.replace(temporary, output_path)
    except (OSError, UnicodeDecodeError, ValueError) as exc:
        return {
            "schema_version": 1,
            "passed": False,
            "contract": kind,
            "output": str(output_path),
            "sha256": None,
            "issues": [f"{kind} schema could not be exported: {exc}"],
        }
    return {
        "schema_version": 1,
        "passed": True,
        "contract": kind,
        "output": str(output_path),
        "sha256": sha256_bytes(encoded),
        "issues": [],
    }
