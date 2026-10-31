# Examples

Run these from the repository root:

    python examples/the_corruption.py    # the paragraph python-docx gets wrong
    python examples/review_notes.py      # comments with anchors, and the revisions no reader reports

    # the switch acceptance test: python-docx's own twelve documented recipes,
    # run through both libraries; the two transcripts diff clean
    python examples/docs_recipes.py                       # with python-docx
    PYTHONPATH=src python examples/docs_recipes.py --shim # with docviews.compat

    # and a program written against the shim, same output either way
    python examples/switch_check.py            # with python-docx
    python examples/switch_check.py --shim     # with docviews.compat

    # authoring: a picture, a header, merged cells, a style, metadata
    python examples/build_report.py out/report.docx
    python examples/build_report.py out/report_shim.docx --shim

    docviews audit . --show 10                 # audit any folder of .docx files
