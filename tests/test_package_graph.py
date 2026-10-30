"""Tests for immutable, reusable PackageGraph."""

from __future__ import annotations

import concurrent.futures
import io
import threading
import zipfile
from dataclasses import FrozenInstanceError
from pathlib import Path
from types import MappingProxyType

import pytest

from wolfdocx.canonical import sha256_bytes, sha256_file
from wolfdocx.inspect import PackageLimits
from wolfdocx.package import (
    PackageFormatError,
    PackageGraph,
    PackageSafetyError,
    PackageXmlElement,
)
from wolfdocx.package_profile import parse_package_profile

SAMPLE_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Default Extension="png" ContentType="image/png"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
    <Override PartName="/customXml/item1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>
    <Override PartName="/word/unknownExtension.dat" ContentType="application/x-proprietary-binary"/>
</Types>
"""

SAMPLE_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.emf"/>
</Relationships>
"""

SAMPLE_DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>
</Relationships>
"""

SAMPLE_DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p w:rsidR="00112233">
            <w:r>
                <w:t>Hello PackageGraph</w:t>
            </w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>
"""


def _create_sample_docx(path: Path, *, extra_parts: dict[str, bytes] | None = None) -> Path:
    parts: dict[str, bytes] = {
        "[Content_Types].xml": SAMPLE_CONTENT_TYPES.encode("utf-8"),
        "_rels/.rels": SAMPLE_ROOT_RELS.encode("utf-8"),
        "word/document.xml": SAMPLE_DOCUMENT_XML.encode("utf-8"),
        "word/_rels/document.xml.rels": SAMPLE_DOC_RELS.encode("utf-8"),
        "word/settings.xml": b'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
        "customXml/item1.xml": b"<customData><value>123</value></customData>",
        "word/unknownExtension.dat": b"\xde\xad\xbe\xef\x01\x02\x03\x04",
        "word/media/image1.png": b"\x89PNG\r\n\x1a\n" + b"\x00" * 2000,
    }
    if extra_parts:
        parts.update(extra_parts)

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return path


def test_package_graph_construction_and_identity(tmp_path: Path) -> None:
    doc_path = _create_sample_docx(tmp_path / "test.docx")
    graph = PackageGraph.from_file(doc_path)

    assert graph.path == doc_path.resolve()
    assert graph.file_sha256 == sha256_file(doc_path)
    assert graph.has_part("word/document.xml")
    assert not graph.has_part("nonexistent.xml")

    doc_part = graph.part("word/document.xml")
    assert doc_part is not None
    assert doc_part.name == "word/document.xml"
    assert (
        doc_part.content_type
        == "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    )

    # Part identity preservation
    assert graph.part("word/document.xml") is doc_part
    assert doc_part.read_bytes() is doc_part.read_bytes()

    # XML projection identity and memoization
    xml_root = doc_part.parsed_xml()
    assert isinstance(xml_root, PackageXmlElement)
    assert doc_part.parsed_xml() is xml_root


def test_package_graph_immutability(tmp_path: Path) -> None:
    doc_path = _create_sample_docx(tmp_path / "test.docx")
    graph = PackageGraph.from_file(doc_path)

    # Graph parts mapping is immutable
    assert isinstance(graph.parts, MappingProxyType)
    with pytest.raises(TypeError):
        graph.parts["new_part"] = None  # type: ignore[index]

    # Graph relationships is an immutable tuple
    assert isinstance(graph.relationships, tuple)
    with pytest.raises((TypeError, AttributeError)):
        graph.relationships.append(None)  # type: ignore[attr-defined]

    # Content types mappings are immutable
    with pytest.raises(TypeError):
        graph.content_type_defaults["extra"] = "val"  # type: ignore[index]
    with pytest.raises(TypeError):
        graph.content_type_overrides["extra"] = "val"  # type: ignore[index]

    # PackageXmlElement is a frozen dataclass
    doc_part = graph.part("word/document.xml")
    assert doc_part is not None
    xml_root = doc_part.parsed_xml()
    with pytest.raises(FrozenInstanceError):
        xml_root.text = "mutated"  # type: ignore[misc]
    with pytest.raises(TypeError):
        xml_root.attrib["key"] = "val"  # type: ignore[index]


def test_package_relationship_resolution(tmp_path: Path) -> None:
    doc_path = _create_sample_docx(tmp_path / "test.docx")
    graph = PackageGraph.from_file(doc_path)

    root_rels = graph.relationships
    assert len(root_rels) == 2
    assert root_rels[0].id == "rId1"
    assert root_rels[0].resolved_target() == "word/document.xml"
    assert not root_rels[0].is_external

    doc_rels = graph.relationships_for("word/document.xml")
    assert len(doc_rels) == 4

    rel_map = {r.id: r for r in doc_rels}
    assert rel_map["rId1"].resolved_target() == "word/settings.xml"
    assert rel_map["rId2"].resolved_target() == "word/media/image1.png"
    assert rel_map["rId3"].is_external
    assert rel_map["rId3"].resolved_target() == "https://example.com"
    assert rel_map["rId4"].resolved_target() == "customXml/item1.xml"


def test_unknown_parts_and_types_preservation(tmp_path: Path) -> None:
    doc_path = _create_sample_docx(tmp_path / "test.docx")
    graph = PackageGraph.from_file(doc_path)

    custom_part = graph.part("customXml/item1.xml")
    assert custom_part is not None
    assert (
        custom_part.content_type
        == "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
    )
    assert custom_part.read_bytes() == b"<customData><value>123</value></customData>"

    unknown_part = graph.part("word/unknownExtension.dat")
    assert unknown_part is not None
    assert unknown_part.content_type == "application/x-proprietary-binary"
    assert unknown_part.read_bytes() == b"\xde\xad\xbe\xef\x01\x02\x03\x04"


def test_lazy_binary_avoidance_and_counters(tmp_path: Path) -> None:
    doc_path = _create_sample_docx(tmp_path / "test.docx")
    graph = PackageGraph.from_file(doc_path)

    # Upon initial construction, ZIP is opened once to inspect headers and topology
    assert graph.counters.zip_open_count == 1

    media_part = graph.part("word/media/image1.png")
    assert media_part is not None
    # Media bytes have NOT been decompressed yet
    assert media_part._cached_bytes is None

    # Decompress media on demand
    media_data = media_part.read_bytes()
    assert media_data.startswith(b"\x89PNG")
    assert media_part._cached_bytes is not None

    # Counters reflect demand reading
    snap = graph.counters.snapshot()
    assert snap["part_bytes_read_count"] >= 3  # [Content_Types], .rels, + image1.png

    # Second read uses cache, no additional zip open or read increment
    current_reads = graph.counters.part_bytes_read_count
    current_opens = graph.counters.zip_open_count
    assert media_part.read_bytes() is media_data
    assert graph.counters.part_bytes_read_count == current_reads
    assert graph.counters.zip_open_count == current_opens


def test_archive_session_reuses_one_archive_without_retaining_xml_bytes(
    tmp_path: Path,
) -> None:
    doc_path = _create_sample_docx(tmp_path / "test.docx")
    graph = PackageGraph.from_file(doc_path)
    document_part = graph.part("word/document.xml")
    settings_part = graph.part("word/settings.xml")
    assert document_part is not None and settings_part is not None

    before = graph.counters.snapshot()
    with graph.archive_session():
        document_part.parsed_xml()
        settings_part.parsed_xml()
    after = graph.counters.snapshot()

    assert after["zip_open_count"] == before["zip_open_count"] + 1
    assert after["current_materialized_bytes"] == before["current_materialized_bytes"]
    assert document_part._cached_bytes is None
    assert settings_part._cached_bytes is None


def test_archive_session_nests_and_closes_after_success(tmp_path: Path) -> None:
    doc_path = _create_sample_docx(tmp_path / "test.docx")
    graph = PackageGraph.from_file(doc_path)
    document_part = graph.part("word/document.xml")
    settings_part = graph.part("word/settings.xml")
    assert document_part is not None and settings_part is not None

    before = graph.counters.snapshot()
    with graph.archive_session(), graph.archive_session():
        document_part.parsed_xml()
        settings_part.parsed_xml()
        archive = graph._archive
        assert archive is not None
        assert graph._archive_session_depth == 2

    assert graph.counters.zip_open_count == before["zip_open_count"] + 1
    assert graph._archive is None
    assert graph._archive_session_depth == 0


def test_archive_session_closes_after_xml_parse_failure(tmp_path: Path) -> None:
    doc_path = _create_sample_docx(
        tmp_path / "test.docx",
        extra_parts={"word/document.xml": b"<w:document"},
    )
    graph = PackageGraph.from_file(doc_path)
    document_part = graph.part("word/document.xml")
    assert document_part is not None

    with (
        pytest.raises(PackageFormatError, match="XML parse error"),
        graph.archive_session(),
    ):
        document_part.parsed_xml()

    assert graph._archive is None
    assert graph._archive_session_depth == 0
    assert graph.counters.current_materialized_bytes == 0
    assert document_part._cached_bytes is None


def test_archive_session_serializes_concurrent_part_reads(tmp_path: Path) -> None:
    doc_path = _create_sample_docx(tmp_path / "test.docx")
    graph = PackageGraph.from_file(doc_path)
    part_names = ("word/document.xml", "word/settings.xml", "customXml/item1.xml")
    barrier = threading.Barrier(len(part_names))
    before = graph.counters.snapshot()

    def parse(part_name: str) -> str:
        part = graph.part(part_name)
        assert part is not None
        with graph.archive_session():
            barrier.wait()
            return part.parsed_xml().tag

    with concurrent.futures.ThreadPoolExecutor(max_workers=len(part_names)) as executor:
        tags = list(executor.map(parse, part_names))

    assert len(tags) == len(part_names)
    assert graph.counters.zip_open_count == before["zip_open_count"] + 1
    assert graph._archive is None
    assert graph._archive_session_depth == 0


def test_concurrent_readers_thread_safety(tmp_path: Path) -> None:
    doc_path = _create_sample_docx(tmp_path / "test.docx")
    graph = PackageGraph.from_file(doc_path)

    def worker(iteration: int) -> tuple[int, int, str]:
        doc_part = graph.part("word/document.xml")
        assert doc_part is not None
        data = doc_part.read_bytes()
        xml = doc_part.parsed_xml()
        body = xml.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body")
        assert body is not None
        p = body.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")
        assert p is not None
        return (len(data), len(xml.children), doc_part.sha256)

    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
        futures = [executor.submit(worker, i) for i in range(32)]
        results = [f.result() for f in concurrent.futures.as_completed(futures)]

    # All threads observe the exact same results
    first = results[0]
    for r in results[1:]:
        assert r == first


def test_immutable_xml_element_semantic_api(tmp_path: Path) -> None:
    doc_path = _create_sample_docx(tmp_path / "test.docx")
    graph = PackageGraph.from_file(doc_path)
    doc_part = graph.part("word/document.xml")
    assert doc_part is not None

    root = doc_part.parsed_xml()
    assert root.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}document"

    # Traversals
    tags = [elem.tag for elem in root.iter()]
    assert "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t" in tags

    t_elements = [
        elem
        for elem in root.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
    ]
    assert len(t_elements) == 1
    assert t_elements[0].text == "Hello PackageGraph"

    # Mutable conversion returns a decoupled copy
    mutable_copy = root.to_mutable_element()
    mutable_copy.attrib["test"] = "123"
    assert "test" not in root.attrib


def test_exact_hash_stability(tmp_path: Path) -> None:
    doc_path = _create_sample_docx(tmp_path / "test.docx")
    graph1 = PackageGraph.from_file(doc_path)
    graph2 = PackageGraph.from_file(doc_path)

    assert graph1.file_sha256 == sha256_file(doc_path)
    assert graph1.file_sha256 == graph2.file_sha256

    doc_part1 = graph1.part("word/document.xml")
    doc_part2 = graph2.part("word/document.xml")
    assert doc_part1 is not None and doc_part2 is not None
    assert doc_part1.sha256 == sha256_bytes(doc_part1.read_bytes())
    assert doc_part1.sha256 == doc_part2.sha256


def test_hostile_zip_path_traversal_rejection(tmp_path: Path) -> None:
    bad_zip = tmp_path / "traversal.docx"
    with zipfile.ZipFile(bad_zip, "w") as zf:
        zf.writestr("[Content_Types].xml", SAMPLE_CONTENT_TYPES)
        zf.writestr("../evil.txt", b"exploit")

    with pytest.raises(PackageSafetyError) as exc_info:
        PackageGraph.from_file(bad_zip)
    assert exc_info.value.code == "unsafe_path"


def test_hostile_zip_absolute_path_rejection(tmp_path: Path) -> None:
    bad_zip = tmp_path / "absolute.docx"
    with zipfile.ZipFile(bad_zip, "w") as zf:
        zf.writestr("[Content_Types].xml", SAMPLE_CONTENT_TYPES)
        zf.writestr("/root_file.xml", b"<xml/>")

    with pytest.raises(PackageSafetyError) as exc_info:
        PackageGraph.from_file(bad_zip)
    assert exc_info.value.code == "unsafe_path"


def test_hostile_zip_backslash_path_rejection(tmp_path: Path) -> None:
    bad_zip = tmp_path / "backslash.docx"
    with zipfile.ZipFile(bad_zip, "w") as zf:
        zf.writestr("[Content_Types].xml", SAMPLE_CONTENT_TYPES)
        zf.writestr("word\\document.xml", b"<xml/>")

    with pytest.raises(PackageSafetyError) as exc_info:
        PackageGraph.from_file(bad_zip)
    assert exc_info.value.code == "unsafe_path"


def test_hostile_zip_nonempty_directory_entry_rejection(tmp_path: Path) -> None:
    bad_zip = tmp_path / "directory.docx"
    with zipfile.ZipFile(bad_zip, "w") as zf:
        zf.writestr("[Content_Types].xml", SAMPLE_CONTENT_TYPES)
        zf.writestr("word/", b"not directory metadata")

    with pytest.raises(PackageSafetyError) as exc_info:
        PackageGraph.from_file(bad_zip)
    assert exc_info.value.code == "directory_entry"


def test_hostile_zip_duplicate_entry_rejection(tmp_path: Path) -> None:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", SAMPLE_CONTENT_TYPES)
        zf.writestr("word/document.xml", b"<w:document/>")
        zf.writestr("word/document.xml", b"<w:document2/>")
    bad_zip = tmp_path / "duplicate.docx"
    bad_zip.write_bytes(buf.getvalue())

    with pytest.raises(PackageFormatError) as exc_info:
        PackageGraph.from_file(bad_zip)
    assert exc_info.value.code == "duplicate_entry"


def test_package_limits_enforcement(tmp_path: Path) -> None:
    doc_path = _create_sample_docx(tmp_path / "test.docx")

    # Max entries limit
    with pytest.raises(PackageSafetyError) as exc:
        PackageGraph.from_file(doc_path, limits=PackageLimits(max_entries=2))
    assert exc.value.code == "entry_limit"

    # Max entry size limit
    with pytest.raises(PackageSafetyError) as exc:
        PackageGraph.from_file(doc_path, limits=PackageLimits(max_entry_size=10))
    assert exc.value.code == "entry_size_limit"

    # Max uncompressed size limit
    with pytest.raises(PackageSafetyError) as exc:
        PackageGraph.from_file(doc_path, limits=PackageLimits(max_uncompressed_size=100))
    assert exc.value.code == "package_size_limit"


def test_missing_required_content_types_rejection(tmp_path: Path) -> None:
    bad_zip = tmp_path / "no_ct.docx"
    with zipfile.ZipFile(bad_zip, "w") as zf:
        zf.writestr("word/document.xml", b"<doc/>")

    with pytest.raises(PackageFormatError) as exc_info:
        PackageGraph.from_file(bad_zip)
    assert exc_info.value.code == "missing_required_part"


def test_not_a_zip_file_rejection(tmp_path: Path) -> None:
    not_zip = tmp_path / "plain.txt"
    not_zip.write_text("just text")

    with pytest.raises(PackageFormatError) as exc_info:
        PackageGraph.from_file(not_zip)
    assert exc_info.value.code == "not_zip"


def _large_xml_part(root: str, words: int = 80_000) -> bytes:
    payload = "".join(f"{number:08x}" for number in range(words))
    return (
        f'<{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:t>{payload}</w:t></{root}>"
    ).encode()


def test_streaming_xml_events_do_not_materialize_large_part(tmp_path: Path) -> None:
    document = _create_sample_docx(
        tmp_path / "streaming.docx",
        extra_parts={"word/document.xml": _large_xml_part("w:document")},
    )
    graph = PackageGraph.from_file(document)
    part = graph.part("word/document.xml")
    assert part is not None

    before = graph.counters.snapshot()
    events = list(part.iter_xml_events())
    feature_counts = graph.wordprocessingml_feature_counts
    after = graph.counters.snapshot()

    assert [event.kind for event in events].count("start") == 2
    assert events[0].tag == (
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}document"
    )
    assert sum(len(event.text or "") for event in events) > 500_000
    assert feature_counts.sections == 0
    assert after["part_bytes_read_count"] == before["part_bytes_read_count"]
    assert after["part_bytes_streamed_count"] == before["part_bytes_streamed_count"] + 2
    assert after["current_materialized_bytes"] == before["current_materialized_bytes"]
    assert after["peak_materialized_bytes"] == before["peak_materialized_bytes"]


def test_profile_enforces_aggregate_retained_byte_bound(tmp_path: Path) -> None:
    document = _create_sample_docx(
        tmp_path / "aggregate-bound.docx",
        extra_parts={
            "word/document.xml": _large_xml_part("w:document"),
            "word/styles.xml": _large_xml_part("w:styles"),
        },
    )
    profile = parse_package_profile(
        {"schema_version": 1, "memory": {"max_materialized_bytes": 1_048_576}}
    )
    graph = PackageGraph.from_file(document, profile=profile)
    document_part = graph.part("word/document.xml")
    styles_part = graph.part("word/styles.xml")
    assert document_part is not None and styles_part is not None

    document_part.read_bytes()

    with pytest.raises(PackageSafetyError) as exc_info:
        styles_part.read_bytes()

    assert exc_info.value.code == "materialization_limit"
    assert graph.counters.peak_materialized_bytes <= profile.memory.max_materialized_bytes
