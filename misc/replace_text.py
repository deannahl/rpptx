def py_replace_text(pres, label, new_text):
    """Replace text in a text box but retain formatting

    Keyword arguments:
    pres -- pptx Pres object containing the text to be replaced
    label -- The unique label of the object containing the text to be replaced
    new_text -- The new text that the paragraph should contain
    """
    # Get shape from label
    target_shape = [
        shape for slide in pres.slides for shape in slide.shapes if shape.name == label
    ]

    if len(target_shape) == 0:
        raise Exception("No object with the specified label was found.")
    elif len(target_shape) > 1:
        raise Exception("More than one shape with that label was found.")

    old_text = target_shape[0]

    paragraph = old_text.text_frame.paragraphs[0]
    p = paragraph._p  # the lxml element containing the `<a:p>` element
    # remove all but the first run
    for idx, run in enumerate(paragraph.runs):
        if idx == 0:
            continue
        p.remove(run._r)
    paragraph.runs[0].text = new_text
