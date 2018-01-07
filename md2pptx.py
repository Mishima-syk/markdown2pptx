from markdown import Markdown
from html.parser import HTMLParser
from pptx import Presentation
import click


class MyHTMLParser(HTMLParser):
    def __init__(self):
        super(MyHTMLParser, self).__init__()
        self.prs = Presentation()
        self.tags = []
        self.slide = ""

    def handle_starttag(self, tag, attrs):
        if tag == "h1":
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[0])
            self.slide = slide
            self.tags.append(tag)
        elif tag == "h2":
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[1])
            self.slide = slide
            self.tags.append(tag)
        elif tag == "h3":
            self.tags.append(tag)
        else:
            print("Encountered a start tag:", tag)

    def handle_endtag(self, tag):
        pass

    def handle_data(self, data):
        if len(self.tags) > 0:
            tag = self.tags.pop()
            if tag == "h1":
                self.slide.shapes.title.text = data
            elif tag == "h2":
                self.slide.shapes.title.text = data
            elif tag == "h3":
                p = self.slide.shapes.placeholders[1].text_frame.add_paragraph()
                p.text = data
            else:
                print("Encountered some data  :", data)

    def close(self, output="sample.pptx"):
        self.prs.save(output)


@click.command()
@click.argument('input')
@click.option('--output', '-o', default='output.pptx', help='Name of output_file.')
def cli(input, output):
    md = Markdown()
    parser = MyHTMLParser()
    md_txt = open(input, "r").read()
    parser.feed(md.convert(md_txt))
    parser.close(output)


if __name__ == '__main__':
    cli()
