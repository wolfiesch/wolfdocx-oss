"""Session table edits through the table operation modules.

Tables, rows, columns, and cells are addressed as the stateless requests address
them: ``table`` is the 1-based table of ``part`` in document order, ``row`` and
``column`` are 1-based within it, and a table is inserted next to a 1-based
``paragraph`` anchor.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import TYPE_CHECKING

from ._bridge import EditHost, SessionEditReport, Staged, bound_fingerprint

if TYPE_CHECKING:
    from ..table_column import TableColumnRequest
    from ..table_delete import TableDeleteRequest
    from ..table_insert import TableInsertRequest
    from ..table_merge import TableMergeRequest
    from ..table_structure import TableRowRequest

_DOCUMENT = "word/document.xml"


class TableEdits(EditHost):
    def insert_table(
        self,
        paragraph: int,
        rows: Sequence[Sequence[str]],
        grid_widths: Sequence[int],
        *,
        position: str = "after",
        style: str | None = None,
        repeat_header: bool = False,
        alignment: str | None = None,
        shading: str | None = None,
        part: str = _DOCUMENT,
    ) -> SessionEditReport:
        """Insert a rectangular table of ``rows`` before or after ``paragraph``."""
        from ..table_insert import (
            TableInsertRequest,
            author_table_insert,
            table_insert_structure_fingerprint,
        )

        def build(staged: Staged) -> TableInsertRequest:
            return TableInsertRequest(
                source=staged.source,
                destination=staged.destination,
                source_sha256=staged.source_sha256,
                target={"paragraph": paragraph, "part": part},
                structure_fingerprint_sha256=bound_fingerprint(
                    table_insert_structure_fingerprint(
                        staged.source, paragraph=paragraph, part=part
                    )
                ),
                rows=rows,
                grid_widths=grid_widths,
                position=position,
                style=style,
                repeat_header=repeat_header,
                alignment=alignment,
                shading=shading,
            )

        return self._apply_operation("table-insert", build, author_table_insert)

    def delete_table(self, table: int, *, part: str = _DOCUMENT) -> SessionEditReport:
        from ..table_delete import (
            TableDeleteRequest,
            author_table_delete,
            table_delete_structure_fingerprint,
        )

        def build(staged: Staged) -> TableDeleteRequest:
            return TableDeleteRequest(
                source=staged.source,
                destination=staged.destination,
                source_sha256=staged.source_sha256,
                target={"table": table, "part": part},
                structure_fingerprint_sha256=bound_fingerprint(
                    table_delete_structure_fingerprint(staged.source, table=table, part=part)
                ),
            )

        return self._apply_operation("table-delete", build, author_table_delete)

    def insert_table_row(
        self,
        table: int,
        row: int,
        *,
        position: str = "after",
        cell_texts: Sequence[str] | None = None,
        repeat_header: bool | None = None,
        cant_split: bool | None = None,
        part: str = _DOCUMENT,
    ) -> SessionEditReport:
        """Insert a row before or after ``row``, cloning its cell structure."""
        return self._row(
            table,
            row,
            part,
            action="insert",
            position=position,
            cell_texts=None if cell_texts is None else tuple(cell_texts),
            repeat_header=repeat_header,
            cant_split=cant_split,
        )

    def delete_table_row(self, table: int, row: int, *, part: str = _DOCUMENT) -> SessionEditReport:
        return self._row(table, row, part, action="delete")

    def insert_table_column(
        self,
        table: int,
        column: int,
        *,
        position: str = "after",
        width: int | None = None,
        cell_texts: Sequence[str] | None = None,
        part: str = _DOCUMENT,
    ) -> SessionEditReport:
        """Insert a column before or after ``column``."""
        return self._column(
            table,
            column,
            part,
            action="insert",
            position=position,
            width=width,
            cell_texts=None if cell_texts is None else tuple(cell_texts),
        )

    def delete_table_column(
        self, table: int, column: int, *, part: str = _DOCUMENT
    ) -> SessionEditReport:
        return self._column(table, column, part, action="delete")

    def merge_table_cells(
        self,
        table: int,
        *,
        start_row: int,
        end_row: int,
        start_col: int,
        end_col: int,
        part: str = _DOCUMENT,
    ) -> SessionEditReport:
        """Merge the rectangle of cells into one."""
        return self._merge(table, start_row, end_row, start_col, end_col, part, "merge")

    def split_table_cells(
        self,
        table: int,
        *,
        start_row: int,
        end_row: int,
        start_col: int,
        end_col: int,
        part: str = _DOCUMENT,
    ) -> SessionEditReport:
        """Split the merged rectangle of cells back into single cells."""
        return self._merge(table, start_row, end_row, start_col, end_col, part, "split")

    def _row(self, table: int, row: int, part: str, **fields: object) -> SessionEditReport:
        from ..table_structure import (
            TableRowRequest,
            author_table_row,
            table_row_structure_fingerprint,
        )

        def build(staged: Staged) -> TableRowRequest:
            return TableRowRequest(
                source=staged.source,
                destination=staged.destination,
                source_sha256=staged.source_sha256,
                target={"table": table, "row": row, "part": part},
                structure_fingerprint_sha256=bound_fingerprint(
                    table_row_structure_fingerprint(staged.source, table=table, row=row, part=part)
                ),
                **fields,  # type: ignore[arg-type]
            )

        return self._apply_operation("table-row", build, author_table_row)

    def _column(self, table: int, column: int, part: str, **fields: object) -> SessionEditReport:
        from ..table_column import (
            TableColumnRequest,
            author_table_column,
            table_column_structure_fingerprint,
        )

        def build(staged: Staged) -> TableColumnRequest:
            return TableColumnRequest(
                source=staged.source,
                destination=staged.destination,
                source_sha256=staged.source_sha256,
                target={"table": table, "column": column, "part": part},
                structure_fingerprint_sha256=bound_fingerprint(
                    table_column_structure_fingerprint(
                        staged.source, table=table, column=column, part=part
                    )
                ),
                **fields,  # type: ignore[arg-type]
            )

        return self._apply_operation("table-column", build, author_table_column)

    def _merge(
        self,
        table: int,
        start_row: int,
        end_row: int,
        start_col: int,
        end_col: int,
        part: str,
        action: str,
    ) -> SessionEditReport:
        from ..table_merge import (
            TableMergeRequest,
            author_table_merge,
            table_merge_structure_fingerprint,
        )

        rectangle = {
            "table": table,
            "start_row": start_row,
            "end_row": end_row,
            "start_col": start_col,
            "end_col": end_col,
        }

        def build(staged: Staged) -> TableMergeRequest:
            return TableMergeRequest(
                source=staged.source,
                destination=staged.destination,
                source_sha256=staged.source_sha256,
                target={**rectangle, "part": part},
                structure_fingerprint_sha256=bound_fingerprint(
                    table_merge_structure_fingerprint(staged.source, part=part, **rectangle)
                ),
                action=action,
            )

        return self._apply_operation("table-merge", build, author_table_merge)
