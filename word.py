import re

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor

document = Document()


def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None


section = document.sections[0]
header = section.header
delete_paragraph(header.paragraphs[0])
table = header.add_table(4, 5, Inches(6))
column = table.columns[0]
column.cells[0].merge(column.cells[-1])
table.style = 'Table Grid'


def writedocx(content='', font_name='Times New Roman', font_size=12, font_bold=False, font_italic=False,
              font_underline=False, color=RGBColor(0, 0, 0),
              before_spacing=0, after_spacing=10, line_spacing=1.5, keep_together=True, keep_with_next=False,
              page_break_before=False,
              widow_control=False, align='left', style='Normal'):
    paragraph = document.add_paragraph(style=style)
    run = paragraph.add_run(str(content))
    print(f'Run Style: {run.style}')
    font = run.font
    font.name = font_name
    font.size = Pt(font_size)
    font.bold = font_bold
    font.italic = font_italic
    font.underline = font_underline
    font.color.rgb = color
    paragraph_format = paragraph.paragraph_format
    paragraph_format.space_before = Pt(before_spacing)
    paragraph_format.space_after = Pt(after_spacing)
    paragraph_format.line_spacing = line_spacing
    paragraph_format.keep_together = keep_together
    paragraph_format.keep_with_next = keep_with_next
    paragraph_format.page_break_before = page_break_before
    paragraph_format.widow_control = widow_control
    if align.lower() == 'left':
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    elif align.lower() == 'center':
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    elif align.lower() == 'right':
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    elif align.lower() == 'justify':
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    else:
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT


with open('test_ho1.txt', 'r', encoding='utf-8') as rf:
    bold = False
    for line in rf:
        x = re.findall(r"\B%[a-zA-Z][a-zA-Z]\b%", line)
        if x and x[0] == '%BA%':
            bold = not bold
            continue
        if bold:
            print(f'BOLD {line}')
            writedocx(content=line, font_bold=True, font_size=17, align='right')
        else:
            print(line)
            writedocx(line)

        # if x:
        #     print(line)
        #     rf.read()
        #     print(line)
        #     writedocx(content=line, style='Body Text 2', font_bold=True)
        # else:
        #     writedocx(line)
        # while not x:
        #     print(line)
        #     break

#
# writedocx(content='Normal', font_size= 15,font_bold=True)
# writedocx(content='Normal', font_size=12)
# writedocx(content='Normal', font_size= 15)
# writedocx(content='Body', style='Body Text')
# writedocx(content='Body', style='Body Text 2')
# writedocx(content='Body', style='Body Text 3')
# writedocx(content='Caption', style='Caption')
# writedocx(content='Heading', style='Heading 1')
# writedocx(content='a', style='Heading 2')
# writedocx(content='a', style='Heading 3')
# writedocx(content='a', style='Heading 4')
# writedocx(content='a', style='Heading 5')
# writedocx(content='a', style='Heading 6')
# writedocx(content='a', style='Heading 7')
# writedocx(content='a', style='Heading 8')
# writedocx(content='a', style='Heading 9')
# writedocx(content='Intense', style='Intense Quote')
# writedocx(content='a', style='List')
# writedocx(content='a', style='List 2')
# writedocx(content='a', style='List 3')
# writedocx(content='a', style='List Bullet')
# writedocx(content='a', style='List Bullet 2')
# writedocx(content='a', style='List Bullet 3')
# writedocx(content='a', style='List Continue')
# writedocx(content='a', style='List Continue 2')
# writedocx(content='a', style='List Continue 3')
# writedocx(content='a', style='List Number')
# writedocx(content='a', style='List Number 2')
# writedocx(content='a', style='List Number 3')
# writedocx(content='a', style='List Paragraph')
# # writedocx(content='a', style='Macro Text')
# writedocx(content='a', style='No Spacing')
# writedocx(content='a', style='Quote')
# writedocx(content='a', style='Subtitle')
# writedocx(content='a', style='TOCHeading')
# writedocx(content='a', style='Title')

document.save('word.docx')
