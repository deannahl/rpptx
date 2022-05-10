def py_replace_table(pres, label, new_table, new_table_shape):
    """Replace text in a text box but retain formatting

    Keyword arguments:
    pres -- pptx Pres object containing the text to be replaced
    label -- The unique label of the object containing the text to be replaced
    new_table -- Tuple containing the contents of the new table elements, in left-to-right,
      top-to-bottom order
    """
    target_shape = [
        shape for slide in pres.slides for shape in slide.shapes if shape.name == label
    ]

    if len(target_shape) == 0:
        raise Exception("No object with the specified label was found.")
    elif len(target_shape) > 1:
        raise Exception("More than one shape with that label was found.")

    old_table = target_shape[0]

    # Check the shape of the old table against the shape of the new table
    if (len(old_table.table.rows) != new_table_shape[0]) | (
        len(old_table.table.columns) != new_table_shape[1]
    ):
        raise ValueError(
            "The number of rows and columns in the new table does not match the old table."
        )

    for table_idx, cell in enumerate(old_table.table.iter_cells()):
        # Replace old text with new text, keeping formatting
        paragraph = cell.text_frame.paragraphs[0]

        for run_idx, run in enumerate(paragraph.runs):
            if run_idx == 0:
                continue
            p.remove(run._r)

        paragraph.runs[0].text = new_table[table_idx]
