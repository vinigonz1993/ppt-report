# ppt-report
This repo consists in using a package called `python-pptx` that generates power pointpresentations using python. However the use of the package is not very simple. For that reason the `ppt-report` makes it easier to generate reports from a python dictiorany using that package.

### Sample
The sample below will generate power point presentation. Every key in the dict represents one slide and their content
```python
    from main import PPTReport

    sample = {
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
        5: {
            "title": {
                "text": "Pie chart with labels and legend",
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
                "chart_type": "PIE"
            }
        },
    }

    report = PPTReport("report", sample)

    report.mount()
```
