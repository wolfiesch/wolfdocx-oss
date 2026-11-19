"""Session picture insert and delete through the picture operation modules."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from .._optional import feature_module
from ..canonical import sha256_bytes
from ._bridge import EditHost, SessionEditReport, Staged, bound_fingerprint, refused

_DOCUMENT = "word/document.xml"


class PictureEdits(EditHost):
    def insert_picture(
        self,
        paragraph: int,
        image: str | Path,
        *,
        extent_cx: int,
        extent_cy: int,
        alt_text: str | None = None,
        anchor: Any | None = None,
        part: str = _DOCUMENT,
    ) -> SessionEditReport:
        """Insert the PNG, JPEG, or GIF at ``image`` into ``paragraph``, sized in EMU."""
        _feature_picture_insert = feature_module("picture_insert")
        PictureInsertRequest = _feature_picture_insert.PictureInsertRequest
        author_picture_insert = _feature_picture_insert.author_picture_insert
        picture_insert_structure_fingerprint = (
            _feature_picture_insert.picture_insert_structure_fingerprint
        )

        image_path = Path(image)
        try:
            image_sha256 = sha256_bytes(image_path.read_bytes())
        except OSError:
            return refused("picture-insert", "picture_insert_image_missing")

        def build(staged: Staged) -> PictureInsertRequest:
            return PictureInsertRequest(
                source=staged.source,
                source_sha256=staged.source_sha256,
                destination=staged.destination,
                paragraph=paragraph,
                structure_fingerprint_sha256=bound_fingerprint(
                    picture_insert_structure_fingerprint(
                        staged.source, paragraph=paragraph, part=part
                    )
                ),
                image_path=image_path,
                image_sha256=image_sha256,
                extent_cx=extent_cx,
                extent_cy=extent_cy,
                alt_text=alt_text,
                anchor=anchor,
                part=part,
            )

        return self._apply_operation("picture-insert", build, author_picture_insert)

    def delete_picture(
        self, paragraph: int, *, drawing: int = 1, part: str = _DOCUMENT
    ) -> SessionEditReport:
        """Delete the ``drawing``-th picture of ``paragraph`` and any media it alone used."""
        _feature_picture_delete = feature_module("picture_delete")
        PictureDeleteRequest = _feature_picture_delete.PictureDeleteRequest
        author_picture_delete = _feature_picture_delete.author_picture_delete
        picture_delete_structure_fingerprint = (
            _feature_picture_delete.picture_delete_structure_fingerprint
        )

        def build(staged: Staged) -> PictureDeleteRequest:
            return PictureDeleteRequest(
                source=staged.source,
                source_sha256=staged.source_sha256,
                destination=staged.destination,
                target={"paragraph": paragraph, "drawing": drawing, "part": part},
                structure_fingerprint_sha256=bound_fingerprint(
                    picture_delete_structure_fingerprint(
                        staged.source, paragraph=paragraph, drawing=drawing, part=part
                    )
                ),
            )

        return self._apply_operation("picture-delete", build, author_picture_delete)
