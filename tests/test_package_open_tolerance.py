"""Package metadata tolerance without relaxing OPC or resource boundaries."""

from __future__ import annotations

import io
import stat
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from wolfdocx import Document
from wolfdocx.inspect import PackageLimits, inspect_docx
from wolfdocx.package import PackageError, PackageFormatError, PackageGraph, PackageSafetyError

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _package(
    path: Path,
    entries: tuple[tuple[str, int, bytes], ...] = (),
    *,
    target: str | None = None,
    override: str | None = None,
    document: bytes | None = None,
) -> Path:
    types = (
        f'<Types xmlns="{CT}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    )
    if override is not None:
        types += f'<Override PartName="/{override}" ContentType="application/octet-stream"/>'
    rels = (
        f'<Relationships xmlns="{REL}"><Relationship Id="main" '
        f'Type="{OFFICE}/officeDocument" Target="word/document.xml"/>'
    )
    if target is not None:
        rels += f'<Relationship Id="extra" Type="{OFFICE}/customXml" Target="{target}"/>'
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.comment = b"preserve package metadata"
        archive.writestr("[Content_Types].xml", types + "</Types>")
        archive.writestr("_rels/.rels", rels + "</Relationships>")
        archive.writestr(
            "word/document.xml",
            document
            or f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>Keep</w:t></w:r></w:p></w:body></w:document>'.encode(),
        )
        for name, mode, payload in entries:
            info = zipfile.ZipInfo(name)
            info.create_system = 3
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = (mode << 16) | (0x10 if name.endswith("/") else 0)
            archive.writestr(info, payload)
    return path


def _assert_metadata_hidden_and_preserved(path: Path, hidden: set[str]) -> None:
    graph = PackageGraph.from_file(path)
    assert hidden.isdisjoint(graph.parts)
    report = inspect_docx(path)
    assert report.valid, [issue.to_dict() for issue in report.issues]
    assert hidden.isdisjoint(part.name for part in report.parts)
    with Document(path) as document:
        assert hidden.isdisjoint(document.part_names)
        for name in hidden:
            with pytest.raises(PackageError) as raised:
                document.part_bytes(name)
            assert raised.value.code == "missing_part"
        output = io.BytesIO()
        saved = document.save(output)
        assert saved.status == "published", saved.code
        assert saved.raw_copied_parts == len(document.part_names)
        assert output.getvalue() == path.read_bytes()


@pytest.mark.parametrize("mode", [stat.S_IFDIR | 0o755, stat.S_IFDIR | 0o775, 0o755, 0])
def test_empty_directory_metadata_is_hidden_and_preserved(tmp_path: Path, mode: int) -> None:
    names = {"word/", "_rels/", "word/media/"}
    path = _package(
        tmp_path / "directories.docx", tuple((name, mode, b"") for name in sorted(names))
    )
    _assert_metadata_hidden_and_preserved(path, names)


@pytest.mark.parametrize(
    "kind", [stat.S_IFLNK, stat.S_IFCHR, stat.S_IFBLK, stat.S_IFIFO, stat.S_IFSOCK]
)
@pytest.mark.parametrize("name", ["word/special", "word/special/", "[trash]/special.dat"])
def test_special_entries_keep_typed_refusal(tmp_path: Path, kind: int, name: str) -> None:
    path = _package(tmp_path / "special.docx", ((name, kind | 0o755, b""),))
    with pytest.raises(PackageSafetyError) as raised:
        PackageGraph.from_file(path)
    assert raised.value.code == "directory_entry"
    assert raised.value.entry_name == name
    report = inspect_docx(path)
    assert not report.valid
    assert "directory_entry" in {issue.code for issue in report.issues}


@pytest.mark.parametrize(
    "name,mode,payload,code",
    [
        ("word/", stat.S_IFDIR | 0o755, b"not empty", "directory_entry"),
        ("word/", stat.S_IFREG | 0o644, b"", "directory_entry"),
        ("word/directory", stat.S_IFDIR | 0o755, b"", "directory_entry"),
        ("../outside/", stat.S_IFDIR | 0o755, b"", "unsafe_path"),
        ("[trash]/../outside.dat", stat.S_IFREG | 0o644, b"", "unsafe_path"),
        ("[trash]/bad%2Fname.dat", stat.S_IFREG | 0o644, b"", "part_name_invalid"),
        ("word/bad[name].xml", stat.S_IFREG | 0o644, b"", "part_name_invalid"),
    ],
)
def test_metadata_exception_does_not_hide_invalid_entries(
    tmp_path: Path, name: str, mode: int, payload: bytes, code: str
) -> None:
    path = _package(tmp_path / "invalid.docx", ((name, mode, payload),))
    with pytest.raises((PackageSafetyError, PackageFormatError)) as raised:
        PackageGraph.from_file(path)
    assert raised.value.code == code
    assert not inspect_docx(path).valid


def test_unreferenced_legacy_trash_is_hidden_and_preserved(tmp_path: Path) -> None:
    names = {"[trash]/", "[trash]/0000.dat", "[trash]/unused.xml"}
    path = _package(
        tmp_path / "trash.docx",
        (
            ("[trash]/", stat.S_IFDIR | 0o755, b""),
            ("[trash]/0000.dat", stat.S_IFREG | 0o644, b"legacy bytes"),
            ("[trash]/unused.xml", stat.S_IFREG | 0o644, b"opaque, not XML"),
        ),
    )
    _assert_metadata_hidden_and_preserved(path, names)


def test_text_edit_preserves_unrelated_parts_and_zip_metadata(tmp_path: Path) -> None:
    path = _package(
        tmp_path / "edit.docx",
        (
            ("word/", stat.S_IFDIR | 0o755, b""),
            ("[trash]/0000.dat", stat.S_IFREG | 0o644, b"legacy bytes"),
            ("word/opaque.dat", stat.S_IFREG | 0o644, b"unrelated bytes"),
        ),
    )

    def raw_members(data: bytes) -> dict[str, bytes]:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            entries = archive.infolist()
            offsets = [entry.header_offset for entry in entries] + [archive.start_dir]
            return {
                entry.filename: data[offsets[index] : offsets[index + 1]]
                for index, entry in enumerate(entries)
            }

    before = raw_members(path.read_bytes())
    with Document(path) as document:
        edited = document.find_replace("Keep", "Changed")
        assert edited.passed, edited.code
        output = io.BytesIO()
        saved = document.save(output)
        assert saved.status == "published", saved.code
    after = raw_members(output.getvalue())
    assert before.keys() == after.keys()
    assert {name for name in before if before[name] != after[name]} == {"word/document.xml"}
    with zipfile.ZipFile(io.BytesIO(output.getvalue())) as archive:
        assert b"Changed" in archive.read("word/document.xml")


def test_package_level_edit_preserves_hidden_metadata(tmp_path: Path) -> None:
    hidden = {"word/": b"", "[trash]/0000.dat": b"legacy bytes"}
    entries = (
        ("word/", stat.S_IFDIR | 0o755, b""),
        ("[trash]/0000.dat", stat.S_IFREG | 0o644, b"legacy bytes"),
    )
    path = _package(tmp_path / "adopt.docx", entries)
    with Document(path) as document:
        styles = ET.Element(f"{{{W}}}styles")
        ET.SubElement(styles, f"{{{W}}}style", {f"{{{W}}}styleId": "Added"})
        document.styles = styles
        assert hidden.keys().isdisjoint(document.part_names)
        assert "word/styles.xml" in document.part_names
        output = io.BytesIO()
        saved = document.save(output)
        assert saved.status == "published", saved.code
    with zipfile.ZipFile(output) as archive:
        assert {name: archive.read(name) for name in hidden} == hidden
        assert b"Added" in archive.read("word/styles.xml")


def test_definition_part_can_be_added_to_plain_document(tmp_path: Path) -> None:
    with Document(_package(tmp_path / "plain.docx")) as document:
        document.lists = ET.Element(f"{{{W}}}numbering")
        assert "word/numbering.xml" in document.part_names


@pytest.mark.parametrize(
    "limits,code",
    [
        (PackageLimits(max_entry_size=1_000), "entry_size_limit"),
        (PackageLimits(max_uncompressed_size=1_000), "package_size_limit"),
        (PackageLimits(max_compression_ratio=20), "compression_ratio_limit"),
    ],
)
def test_ignored_trash_still_counts_toward_zip_limits(
    tmp_path: Path, limits: PackageLimits, code: str
) -> None:
    path = _package(
        tmp_path / "trash-limit.docx",
        (("[trash]/0000.dat", stat.S_IFREG, b"x" * 2_000),),
    )
    with pytest.raises(PackageSafetyError) as raised:
        PackageGraph.from_file(path, limits=limits)
    assert raised.value.code == code
    assert raised.value.entry_name == (None if code == "package_size_limit" else "[trash]/0000.dat")
    report = inspect_docx(path, limits=limits)
    assert code in {issue.code for issue in report.issues}


@pytest.mark.parametrize(
    "target", ["[trash]/0000.dat", "/[trash]/0000.dat", "%5Btrash%5D/0000.dat"]
)
def test_referenced_legacy_trash_is_refused(tmp_path: Path, target: str) -> None:
    path = _package(
        tmp_path / "referenced.docx",
        (("[trash]/0000.dat", stat.S_IFREG | 0o644, b"legacy bytes"),),
        target=target,
    )
    with pytest.raises(PackageFormatError) as raised:
        PackageGraph.from_file(path)
    assert raised.value.code == "part_name_invalid"
    assert not inspect_docx(path).valid


@pytest.mark.parametrize(
    "override", ["[trash]/0000.dat", "word/bad[name].xml", "word/bad%2Fname.xml"]
)
def test_invalid_override_name_is_refused(tmp_path: Path, override: str) -> None:
    path = _package(tmp_path / "override.docx", override=override)
    with pytest.raises(PackageFormatError) as raised:
        PackageGraph.from_file(path)
    assert raised.value.code == "part_name_invalid"
    assert not inspect_docx(path).valid


@pytest.mark.parametrize("prefix", ["word/", "[trash]/"])
def test_equivalent_names_still_refuse(tmp_path: Path, prefix: str) -> None:
    path = _package(
        tmp_path / "equivalent.docx",
        ((prefix + "same.dat", stat.S_IFREG, b"a"), (prefix + "SAME.dat", stat.S_IFREG, b"b")),
    )
    with pytest.raises(PackageFormatError) as raised:
        PackageGraph.from_file(path)
    assert raised.value.code == "part_name_equivalent"


def test_default_limits_accept_four_hundred_thousand_elements(tmp_path: Path) -> None:
    # Vary text to keep the synthetic fixture below the unchanged compression limit.
    xml = (
        f'<w:document xmlns:w="{W}"><w:body><w:p>'.encode()
        + b"".join(f"<w:r><w:t>{index}</w:t></w:r>".encode() for index in range(200_000))
        + b"</w:p></w:body></w:document>"
    )
    path = _package(tmp_path / "large.docx", document=xml)
    graph = PackageGraph.from_file(path)
    assert sum(1 for _ in graph.parse_part_xml("word/document.xml").iter()) == 400_003
    with Document(path) as document:
        output = io.BytesIO()
        saved = document.save(output)
        assert saved.status == "published", saved.code
        assert output.getvalue() == path.read_bytes()


@pytest.mark.parametrize(
    "limits,code",
    [
        (PackageLimits(max_total_xml_elements=2), "xml_total_element_limit"),
        (PackageLimits(max_xml_part_size=2), "xml_size_limit"),
        (PackageLimits(max_xml_depth=2), "xml_depth_limit"),
        (PackageLimits(max_xml_elements=2), "xml_element_limit"),
        (PackageLimits(max_entry_size=2), "entry_size_limit"),
        (PackageLimits(max_uncompressed_size=2), "package_size_limit"),
    ],
)
def test_explicit_limits_still_refuse(tmp_path: Path, limits: PackageLimits, code: str) -> None:
    path = _package(tmp_path / "bounded.docx")
    with pytest.raises(PackageSafetyError) as raised:
        graph = PackageGraph.from_file(path, limits=limits)
        graph.parse_part_xml("word/document.xml")
    assert raised.value.code == code
