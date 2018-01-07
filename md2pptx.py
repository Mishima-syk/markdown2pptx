from markdown import Markdown
from html.parser import HTMLParser
from pptx import Presentation
from pptx.util import Inches, Pt
import click
from logging import getLogger, StreamHandler, DEBUG
logger = getLogger(__name__)
handler = StreamHandler()
handler.setLevel(DEBUG)
logger.setLevel(DEBUG)
logger.addHandler(handler)
logger.propagate = False


class MyHTMLParser(HTMLParser):
    def __init__(self):
        super(MyHTMLParser, self).__init__()
        self.prs = Presentation()
        self.tags = []
        self.slide = None  # current slide
        self.ln = 0        # layout number

    def handle_starttag(self, tag, attrs):
        attr_dict = dict(attrs)
        if tag == "h1":
            ln = int(attr_dict.get("class", "0"))
            self.ln = ln
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[ln])
            self.slide = slide
            self.tags.append(tag)
        elif tag == "h2":
            ln = int(attr_dict.get("class", "1"))
            self.ln = ln
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[ln])
            self.slide = slide
            self.tags.append(tag)
        elif tag == "img":
            img_path = attr_dict.get("src", None)
            left = top = Inches(1)
            self.slide.shapes.add_picture(img_path, left, top)
        else:
            self.tags.append(tag)
        logger.debug(self.tags)

    def handle_endtag(self, tag):
        pass

    def handle_data(self, data):
        if self.ln == 0:
            self.handle_data_layout0(data)
        elif self.ln == 1:
            self.handle_data_layout1(data)
        elif self.ln == 2:
            self.handle_data_layout2(data)
        elif self.ln == 3:
            self.handle_data_layout3(data)
        elif self.ln == 4:
            self.handle_data_layout4(data)
        elif self.ln == 5:
            self.handle_data_layout5(data)
        elif self.ln == 6:
            self.handle_data_layout6(data)
        elif self.ln == 7:
            self.handle_data_layout7(data)
        elif self.ln == 8:
            self.handle_data_layout8(data)
        elif self.ln == 9:
            self.handle_data_layout9(data)
        else:
            print("Not Implemented...")

    def handle_data_layout0(self, data):
        tag = self.tags.pop()
        if tag == "h1":
                self.slide.shapes.title.text = data
        else:
            print("Not Implemented...")

    def handle_data_layout1(self, data):
        tag = self.tags.pop()
        if tag == "h2":
            self.slide.shapes.title.text = data
        elif tag == "h3":
            p = self.slide.shapes.placeholders[1].text_frame.add_paragraph()
            p.text = data
        elif tag == "h4":
            p = self.slide.shapes.placeholders[1].text_frame.add_paragraph()
            p.text = data
            p.level = 1
        elif tag == "h5":
            p = self.slide.shapes.placeholders[1].text_frame.add_paragraph()
            p.text = data
            p.level = 2
        elif tag == "h6":
            p = self.slide.shapes.placeholders[1].text_frame.add_paragraph()
            p.text = data
            p.level = 3
        else:
            print("Not Implemented...")

    def handle_data_layout2(self, data):
        tag = self.tags.pop()
        print("Not Implemented...")

    def handle_data_layout3(self, data):
        tag = self.tags.pop()
        print("Not Implemented...")

    def handle_data_layout4(self, data):
        tag = self.tags.pop()
        print("Not Implemented...")

    def handle_data_layout5(self, data):
        tag = self.tags.pop()
        print("Not Implemented...")

    def handle_data_layout6(self, data):
        tag = self.tags.pop()
        if tag == "h2":
            pass
        elif tag == "p":
            tb = self.slide.shapes.add_textbox(Inches(1), Inches(1), Inches(1), Inches(1))
            tb.text_frame.text = data
        else:
            print("Not Implemented...")

    def handle_data_layout7(self, data):
        tag = self.tags.pop()
        print("Not Implemented...")

    def handle_data_layout8(self, data):
        tag = self.tags.pop()
        print("Not Implemented...")

    def handle_data_layout9(self, data):
        tag = self.tags.pop()
        print("Not Implemented...")

    def close(self, output="sample.pptx"):
        self.prs.save(output)


@click.command()
@click.argument('input')
@click.option('--output', '-o', default='output.pptx', help='Name of output_file.')
def cli(input, output):
    md = Markdown(extensions=['markdown.extensions.attr_list'])
    parser = MyHTMLParser()
    md_txt = open(input, "r").read()
    parser.feed(md.convert(md_txt).replace("\n", ""))  # ???
    parser.close(output)


if __name__ == '__main__':
    cli()
