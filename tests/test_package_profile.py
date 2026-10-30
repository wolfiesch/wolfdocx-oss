from __future__ import annotations

import pytest

from tests.docx_fixtures import paragraph, write_package
from wolfdocx import PackageLimits
from wolfdocx.inspect import inspect_docx
from wolfdocx.package_profile import (
    MAX_MAX_MATERIALIZED_BYTES,
    MIN_MAX_MATERIALIZED_BYTES,
    PackageProfileError,
    load_package_profile,
    parse_package_profile,
)


def _profile(max_materialized_bytes: int = 1024 * 1024) -> dict[str, object]:
    return {
        "schema_version": 1,
        "memory": {"max_materialized_bytes": max_materialized_bytes},
    }


@pytest.mark.parametrize(
    ("value", "code"),
    [
        ([], "profile_not_object"),
        (
            {
                "schema_version": 1,
                "memory": {"max_materialized_bytes": 1024},
                "unknown": True,
            },
            "profile_unknown_field",
        ),
        (
            {"schema_version": "1", "memory": {"max_materialized_bytes": 1024}},
            "profile_schema_version_invalid",
        ),
        ({"schema_version": 1, "memory": []}, "profile_memory_not_object"),
        (
            {"schema_version": 1, "memory": {"unknown": 1024}},
            "profile_memory_unknown_field",
        ),
        (
            {"schema_version": 1, "memory": {"max_materialized_bytes": True}},
            "profile_memory_max_materialized_bytes_invalid",
        ),
        (
            _profile(MIN_MAX_MATERIALIZED_BYTES - 1),
            "profile_memory_max_materialized_bytes_out_of_range",
        ),
        (
            _profile(MAX_MAX_MATERIALIZED_BYTES + 1),
            "profile_memory_max_materialized_bytes_out_of_range",
        ),
    ],
)
def test_profile_parser_rejects_invalid_values_deterministically(value, code):
    with pytest.raises(PackageProfileError) as first:
        parse_package_profile(value)
    with pytest.raises(PackageProfileError) as second:
        parse_package_profile(value)

    assert first.value.code == code
    assert second.value.code == code
    assert str(first.value) == code


def test_profile_loader_rejects_duplicate_json_keys(tmp_path):
    profile_path = tmp_path / "profile.json"
    profile_path.write_text(
        '{"schema_version":1,"schema_version":1,"memory":{"max_materialized_bytes":1024}}',
        encoding="utf-8",
    )

    with pytest.raises(PackageProfileError) as exc_info:
        load_package_profile(profile_path)

    assert exc_info.value.code == "profile_invalid_json"


def test_profile_is_exact_round_trippable_and_exposes_memory_bound(tmp_path):
    profile_path = tmp_path / "profile.json"
    profile_path.write_text(
        '{"schema_version":1,"memory":{"max_materialized_bytes":1048576}}',
        encoding="utf-8",
    )

    profile = load_package_profile(profile_path)

    assert profile.memory.max_materialized_bytes == 1_048_576
    assert profile.to_dict() == _profile(1_048_576)


def test_profile_memory_bound_is_consumed_by_package_inspection(tmp_path):
    # The body text alone exceeds the minimum materialized-bytes bound.
    document = write_package(tmp_path / "bounded.docx", paragraph("Bounded body text. " * 100))
    profile = parse_package_profile(_profile(MIN_MAX_MATERIALIZED_BYTES))

    report = inspect_docx(document, profile=profile)

    assert report.valid is False
    assert [issue.code for issue in report.issues] == ["package_size_limit"]
    assert report.parts == []
    assert report.limits["max_entry_size"] == MIN_MAX_MATERIALIZED_BYTES
    assert report.limits["max_uncompressed_size"] == MIN_MAX_MATERIALIZED_BYTES
    assert report.limits["max_xml_part_size"] == MIN_MAX_MATERIALIZED_BYTES


def test_default_package_limits_remain_explicitly_bounded():
    limits = PackageLimits()

    assert limits.max_archive_size > 0
    assert limits.max_entry_size > 0
    assert limits.max_uncompressed_size > 0
    assert limits.max_xml_part_size > 0
