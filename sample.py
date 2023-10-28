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
                "text": "Second",
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
                "text": "Third slide of the presentation",
                **title_props
            },
            "bar_chart": {
                "title": "Sample Chart",
                "data": [
                    ["2022", [10, 20, 30, 35]],
                    ["2023", [13, 21, 31, 25]]
                ],
                "categories": ["Apple", "Google", "Amazon", "Microsoft"]
            }
        },
    }
