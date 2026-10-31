"""Reproducible build-and-save comparison for the documented authoring path.

Run from a checkout with the optional comparison dependencies installed:

    PYTHONPATH=src python tools/bench_compat_write.py

The timer covers construction of N body paragraphs plus a populated 20x4 table,
then writing to an in-memory stream. Every output is reopened with both readers
and checked for the same paragraph and table content. Results are medians; this
is a diagnostic microbenchmark, not a universal speed claim.
"""
from __future__ import annotations

import io
import platform
import statistics
import sys
import time
from typing import Callable

import docx
from docviews.compat import Document as CompatDocument


def build_and_save(factory: Callable, paragraph_count: int) -> tuple[float, float, bytes]:
    document = factory()
    started = time.perf_counter()
    for index in range(paragraph_count):
        document.add_paragraph(f"paragraph {index}: compatibility benchmark")
    table = document.add_table(rows=20, cols=4)
    for row in range(20):
        for column in range(4):
            table.cell(row, column).text = f"{row}-{column}"
    build_seconds = time.perf_counter() - started

    stream = io.BytesIO()
    started = time.perf_counter()
    document.save(stream)
    save_seconds = time.perf_counter() - started
    return build_seconds, save_seconds, stream.getvalue()


def verify(data: bytes, paragraph_count: int) -> None:
    """Check both libraries see the same content the benchmark created."""
    for loader, source in ((docx.Document, io.BytesIO(data)),
                           (CompatDocument, data)):
        document = loader(source)
        assert len(document.paragraphs) == paragraph_count
        assert [paragraph.text for paragraph in document.paragraphs] == [
            f"paragraph {index}: compatibility benchmark"
            for index in range(paragraph_count)
        ]
        assert len(document.tables) == 1
        assert [[document.tables[0].cell(row, column).text
                 for column in range(4)] for row in range(20)] == [
            [f"{row}-{column}" for column in range(4)]
            for row in range(20)
        ]


def measure(factory: Callable, paragraph_count: int, rounds: int = 5):
    observations = [build_and_save(factory, paragraph_count)
                   for _ in range(rounds)]
    for _, _, package in observations:
        verify(package, paragraph_count)
    build = statistics.median(row[0] for row in observations)
    save = statistics.median(row[1] for row in observations)
    size = len(observations[-1][2])
    return build, save, size


def main() -> None:
    print(f"Python {platform.python_version()} | python-docx {docx.__version__} | "
          f"{platform.platform()}")
    print("Median of 5 runs; build = paragraphs + populated 20x4 table; "
          "save = in-memory stream")
    for count in (500, 5000):
        print(f"\n{count:,} paragraphs")
        for label, factory in (("python-docx", docx.Document),
                               ("docviews.compat", CompatDocument)):
            build, save, size = measure(factory, count)
            print(f"  {label:16s} build {build:.3f}s  save {save:.3f}s  "
                  f"total {build + save:.3f}s  package {size:,} B")


if __name__ == "__main__":
    main()
