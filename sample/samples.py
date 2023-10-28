from pptx.util import Inches


def sample():
    title_props = {
        "top": Inches(1),
        "left": Inches(0.5),
        "width": Inches(5)
    }

    return {
        1: {
            "title": {
                "text": "Title 1",
                **title_props
            },
            "text": {
                "text": "Sample text 1"
            }
        },
        2: {
            "title": {
                "text": "Simple bar chart",
                **title_props
            },
            "bar_chart": {
                "title": "Sample Chart",
                "data": [
                    ["Series 1", [10, 20, 30, 35]]
                ],
                "categories": ["Apple", "Google", "Amazon", "Microsoft"]
            }
        },
        3: {
            "title": {
                "text": "Bar chart with labels",
                **title_props
            },
            "bar_chart": {
                "title": "Sample Chart",
                "data": [
                    ["2022", [10, 20, 30, 35]],
                    ["2023", [13, 21, 31, 25]]
                ],
                "categories": ["Apple", "Google", "Amazon", "Microsoft"],
                "labels": True
            }
        },
        4: {
            "title": {
                "text": "Bar chart with labels and legend",
                **title_props
            },
            "bar_chart": {
                "title": "Sample Chart",
                "data": [
                    ["2022", [10, 20, 30, 35]],
                    ["2023", [13, 21, 31, 25]]
                ],
                "categories": ["Apple", "Google", "Amazon", "Microsoft"],
                "labels": True,
                "legend": True
            }
        },
        4: {
            "title": {
                "text": "Line chart with labels and legend",
                **title_props
            },
            "bar_chart": {
                "title": "Sample Chart",
                "data": [
                    ["2022", [10, 20, 30, 35]],
                    ["2023", [13, 21, 31, 25]]
                ],
                "categories": ["Apple", "Google", "Amazon", "Microsoft"],
                "labels": True,
                "legend": True,
                "chart_type": "LINE"
            }
        },
    }
