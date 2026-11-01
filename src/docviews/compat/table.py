"""python-docx-shaped table, row, column, and cell proxies."""

from . import (_CompatCell, _CompatColumn, _CompatColumns, _CompatRow,
               _CompatRows, _CompatTable)

Table = _CompatTable
_Cell = _CompatCell
_Column = _CompatColumn
_Columns = _CompatColumns
_Row = _CompatRow
_Rows = _CompatRows

__all__ = ["Table", "_Cell", "_Column", "_Columns", "_Row", "_Rows"]
