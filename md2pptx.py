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
        if tag == "h1" or tag == "h2":
            hln = int(tag[-1]) - 1
            ln = int(attr_dict.get("class", hln))
            self.ln = ln
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[ln])
            self.slide = slide
            self.tags.append(tag)
        else:
            self.tags.append(tag)
        getattr(self, "handle_starttag_layout{}".format(self.ln))(tag, attr_dict)

    def handle_starttag_layout0(self, tag, attrs):
            pass

    def handle_starttag_layout1(self, tag, attrs):
        if tag == "h2":
            pass
        else:
            self.tags.append(tag)

    def handle_starttag_layout2(self, tag, attr_dict):
            print("Not Implemented...")

    def handle_starttag_layout3(self, tag, attr_dict):
            print("Not Implemented...")

    def handle_starttag_layout4(self, tag, attr_sict):
            print("Not Implemented...")

    def handle_starttag_layout5(self, tag, attr_dict):
            print("Not Implemented...")

    def handle_starttag_layout6(self, tag, attr_dict):
        if tag == "img":
            img_path = attr_dict.get("src", None)
            left = top = Inches(1)
            self.slide.shapes.add_picture(img_path, left, top)
        else:
            self.tags.append(tag)

    def handle_starttag_layout7(self, tag, attr_dict):
            print("Not Implemented...")

    def handle_starttag_layout8(self, tag, attr_dict):
            print("Not Implemented...")

    def handle_starttag_layout9(self, tag, attr_dict):
            print("Not Implemented...")

    def handle_endtag(self, tag):
        pass

    def handle_data(self, data):
        getattr(self, "handle_data_layout{}".format(self.ln))(data)

    def handle_data_layout0(self, data):
        tag = self.tags.pop()
        if tag == "h1":
                self.slide.shapes.title.text = data
        elif tag == "p":
            self.slide.placeholders[1].text = data
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
