import win32com.client
from dataclasses import dataclass

import wordconsts as WC
from const_parse_doc import *


@dataclass
class Margin:
    _value: float
    _fields: str
    _displayNames: str


results = []


def processStep(actual, expected, match, desc):
    if not match:
        results.append([actual, expected, desc])


word = win32com.client.Dispatch("Word.Application")
path = r"C:\Users\sasha\Documents\2.docx"
doc = word.Documents.Open(path)
results = []
pageSetup = doc.Sections[0].PageSetup
margin = [
    Margin(56.7, "BottomMargin", "снизу"),
    Margin(56.7, "TopMargin", "сверху"),
    Margin(85.05, "LeftMargin", "слева"),
    Margin(28.3, "RightMargin", "справа"),
]

# Проверки нумерации страниц
footers = doc.Sections(1).Footers(1)

showFirstPageNumber = footers.PageNumbers.ShowFirstPageNumber
startingNumber = footers.PageNumbers.StartingNumber
alignCenterForFooter = footers.Range.ParagraphFormat.Alignment == WC.wdAlignPageNumberCenter

if showFirstPageNumber:
    print("На первой странице не должно быть номера страницы")
if startingNumber != 0:
    print("Нумерация страниц должна начинаться с 1")
if alignCenterForFooter != 1:
    print("Выравнивание нумерации страниц должно быть по центру")
# Конец проверки нумерации страниц

def pc(f, s, p):  # precision compare
    return abs(f - s) <= p


for i in margin:
    actual = round(pageSetup.__getattr__(i._fields), 2)
    expected = i._value
    name = "Поле " + i._displayNames
    processStep(actual, expected, pc(actual, expected, 0.5), name)

for i in range(len(doc.Paragraphs)):
    p = doc.Paragraphs[i]
    r = p.Range
    textRange = r.Text
    font = r.Font
    if not r.Text.strip():
        continue

    processStep(font.Name, MAIN_FONT, font.Name == MAIN_FONT, f"Шрифт (семейство) {i}")
    processStep(font.Size, SIZE_FONT, font.Size == SIZE_FONT, f"Размер шрифта {i}")
    pf = r.ParagraphFormat
    spacebefore = spaceafter = 0

    processStep(pf.Alignment, WC.wdAlignParagraphJustify, pf.Alignment == WC.wdAlignParagraphJustify, f"Выравнивание {i}")
    processStep(pf.SpaceBefore, spacebefore, pf.SpaceBefore == spacebefore, f"Интервал перед абзацем {i}")
    processStep(pf.SpaceAfter, spaceafter, pf.SpaceAfter == spaceafter, f"Интервал после абзаца {i}")
    processStep(pf.LineSpacingRule, WC.wdLineSpace1pt5, pf.LineSpacingRule == WC.wdLineSpace1pt5, f"Междустрочное расстояние {i}")
    processStep(pf.FirstLineIndent, FIRST_LINE_INDENT, pc(pf.FirstLineIndent, FIRST_LINE_INDENT, 0.25), f"Отступ абзацной строки {i}")
    processStep(pf.Hyphenation, True, pf.Hyphenation, f"Автоматические переносы {i}")

prevPage = 1

# for index_paragraphs in range(len(doc.Paragraphs)):
#     paragraph = doc.Paragraphs[index_paragraphs]
#     r = paragraph.Range.Information(WC.wdActiveEndPageNumber)
#     paragraphFormat = paragraph.ParagraphFormat
#     if r == 1: continue
#     if r != prevPage:
#         if paragraph.Range.Text.strip() and paragraphFormat.Bold and prevPage != r:
#             prevPage = r
#             processStep(paragraphFormat.SpaceBefore, SPACE_BEFORE_HEADERS,
#                         pc(paragraphFormat.SpaceBefore, SPACE_BEFORE_HEADERS, 0.3),
#                         f"Отступ снизу после заголовка {index_paragraphs}")
#
#     print(r)

pgs = doc.ActiveWindow.Panes(1).Pages
cnt = pgs.Count
minimal = 40 + 3
processStep(cnt, minimal, cnt >= minimal, "Объем документа (предварительный)")

for res in results:
    print(res)
