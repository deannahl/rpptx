from pptx.chart.data import CategoryChartData


def py_replace_category_plot(pres, label, categories, series_levels, series_values):
    """Replace the data underlying a category plot with the target label

    Args:
        pres (Presentation): A pptx Presentation object
        label (string): The target label
        categories (list): A list containing the levels of the category factor (generally the x-axis variable)
        series_list (list): A list containing list elements of the form `['SeriesName', (1, 2, 3)]`
    """
    # Find shape from label
    target_shape = [
        shape for slide in pres.slides for shape in slide.shapes if shape.name == label
    ]

    if len(target_shape) == 0:
        raise Exception("No object with the specified label was found.")
    elif len(target_shape) > 1:
        raise Exception("More than one shape with that label was found.")

    old_shape = target_shape[0]

    new_data = CategoryChartData()
    new_data.categories = categories

    for idx, series in enumerate(series_values):
        if not isinstance(series, list):
            series = [series]
        if not isinstance(series_levels, list):
            series_levels = [series_levels]
        new_data.add_series(series_levels[idx], tuple(series))

    old_plot.chart.replace_data(new_data)
