"""The one-paragraph demo: two real views, and the string Word cannot display.

Uses the docx-revisions fixture from the census corpus, whose own text states
what the accepted and rejected readings are.
"""

from pathlib import Path

from docviews import open_docx

CORPUS = Path(__file__).resolve().parent.parent / "corpus"
FIXTURE = CORPUS / "docx-revisions-mai_tests_OXML_TrackChanges_Test.docx"


def main() -> int:
    doc = open_docx(FIXTURE)
    marker = "We prefer"
    for view in ("accepted", "original"):
        for p in doc.paragraphs(view):
            if marker in p.text:
                print(f"{view:<9}: {p.text.split(' [Test')[0]!r}")

    print("\nthe same paragraph through python-docx:")
    import docx
    for p in docx.Document(str(FIXTURE)).paragraphs:
        if marker in p.text:
            print(f"{'reader':<9}: {p.text.split(' [Test')[0]!r}")

    print("\nrevisions docviews reports:")
    for r in doc.revisions():
        print(f"  {r.kind:<10} {r.author:<12} {r.date[:10]}  paragraph {r.paragraph_index}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
