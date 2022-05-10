def get_shape_with_label(pres, label, slide_num=None):
    """Select the shape with a particular label from a presentation object.

    Keyword Arguments:
      pres -- pptx Pres object containing the target shape
      slide_num -- (Optional) Slide number containing the target shape (if known)
      label -- The unique label of the target shape
    """
    if slide_num is None:
        target_shape = [
            shape
            for slide in pres.slides
            for shape in slide.shapes
            if shape.name == label
        ]
    else:
        slide_num = int(slide_num)
        slide = pres.slides[slide_num - 1]
        target_shape = [shape for shape in slide.shapes if shape.name == label]

    if len(target_shape) == 0:
        raise Exception("No object with the specified label was found.")
    elif len(target_shape) > 1:
        raise Exception("More than one shape with that label was found.")

    return target_shape[0]
