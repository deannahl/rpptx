def get_slide_index_with_label(pres, label):
    """Find the (zero-indexed) index of the slide in a Presentation object
        containing an object with the target label.

    Args:
        pres (Presentation): A pptx Presentation object
        label (string): The target label

    Returns:
        numeric: Index of the slide containing an object with the target label
    """
    slide_idx_list = []
    for slide_idx, slide in enumerate(pres.slides):
        target_shape = [shape for shape in slide.shapes if shape.name == label]
        slide_contains_target_shape = len(target_shape) > 0

        if slide_contains_target_shape:
            slide_idx_list.append(slide_idx)

    if len(slide_idx_list) == 1:
        return slide_idx_list[0]
    elif len(slide_idx_list) == 0:
        raise Exception("No object with the specified label was found.")
    elif len(slide_idx_list) > 1:
        raise Exception("More than one object with the specified label was found.")
