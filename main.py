from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches
from sample import sample


class PPTReport:
    '''
        Class to manage the creation of a PPT presentation.
        It is possible to easily create a presentation using
        a dictionary of slides and slide properties.
    '''

    def __init__(self, name, properties):
        self.presentation = Presentation()
        self.layout = self.presentation.slide_layouts[0]
        self.name = name
        self.properties = properties

    def add_slide(self, properties):
        '''
            Adds a slide
        '''
        slide = self.presentation.slides.add_slide(self.layout)

        if "title" in properties:
            title = slide.shapes.title

            for attr in properties["title"]:
                setattr(title, attr, properties["title"][attr])

        if "text" in properties:
            text = slide.placeholders[1]

            for attr in properties["text"]:
                setattr(text, attr, properties["text"][attr])

        if "bar_chart" in properties:
            bar_chart = properties["bar_chart"]
            self.add_bar_chart(
                slide,
                bar_chart.get("title"),
                bar_chart.get("data", []),
                bar_chart.get("categories", [])
            )

        return slide

    def add_bar_chart(self, slide, chart_title="", data=[], categories=[]):
        '''
            Adds a chart to the specific slide
        '''

        if not data:
            print("Categories are missing")
            return

        if not categories:
            print("Categories are missing")
            return

        chart_data = CategoryChartData()
        chart_data.title = chart_title
        chart_data.categories = categories

        for serie in data:
            print(serie)
            chart_data.add_series(
                serie[0],
                serie[1]
            )

        x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4.5)
        slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
        )

    def mount(self):
        '''Mounts the presentation'''
        for i in self.properties.keys():
            report.add_slide(properties=self.properties[i])

        self.presentation.save(f"{self.name}.pptx")


report = PPTReport("report", sample())

report.mount()
