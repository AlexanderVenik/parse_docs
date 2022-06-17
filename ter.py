import win32com.client
from dataclasses import dataclass

import wordconsts as WC
from const_parse_doc import *
from Error import Error


def isParagraphCorrespond(in_paragraph) -> bool:
    """
    Текст соотвествует ГОСТу
    :param in_paragraph: параграф
    """
    return in_paragraph.Font.Size == SIZE_FONT and in_paragraph.Font.Name == MAIN_FONT


def isHeading(in_paragraph) -> bool:
    """
    Проверка, что текст заголовок
    :param in_paragraph: параграф
    """
    return isParagraphCorrespond(in_paragraph) and in_paragraph.Text.isupper() and in_paragraph.Bold == BOLD_TEXT


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
doc = word.Documents.Open(PATH_DOCX)
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
if startingNumber != START_NUMBER:
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
    active_end_page_number = p.Range
    text_range = active_end_page_number.Text
    font = active_end_page_number.Font
    if not active_end_page_number.Text.strip() and not isHeading(active_end_page_number):
        continue

    processStep(font.Name, MAIN_FONT, font.Name == MAIN_FONT, f"Шрифт (семейство) {i}")
    processStep(font.Size, SIZE_FONT, font.Size == SIZE_FONT, f"Размер шрифта {i}")
    pf = active_end_page_number.ParagraphFormat
    space_before = space_after = 0

    if isHeading(active_end_page_number):
        continue
    else:
        processStep(pf.Alignment, WC.wdAlignParagraphJustify, pf.Alignment == WC.wdAlignParagraphJustify, f"Выравнивание {i}")

    # processStep(pf.SpaceBefore, space_before, pf.SpaceBefore == space_before, f"Интервал перед абзацем {i}")
    # processStep(pf.SpaceAfter, space_after, pf.SpaceAfter == space_after, f"Интервал после абзаца {i}")
    processStep(pf.LineSpacingRule, WC.wdLineSpace1pt5, pf.LineSpacingRule == WC.wdLineSpace1pt5,
                f"Междустрочное расстояние {i}")
    processStep(pf.FirstLineIndent, FIRST_LINE_INDENT, pc(pf.FirstLineIndent, FIRST_LINE_INDENT, 0.25),
                f"Отступ абзацной строки {i}")
    processStep(pf.Hyphenation, True, pf.Hyphenation, f"Автоматические переносы {i}")

# Проверка кол-во страниц
pgs = doc.ActiveWindow.Panes(1).Pages
cntPages = pgs.Count
processStep(cntPages, MINIMUM_PAGE, cntPages >= MINIMUM_PAGE, "Объем документа (предварительный)")

prevPage = 1

active_end_page_number = 1
idx_par = 1

if cntPages > 1:
    for index_paragraph in range(len(doc.Paragraphs)):
        number_page = doc.Paragraphs[index_paragraph].Range.Information(WC.wdActiveEndPageNumber)
        if number_page > 2:
            active_end_page_number = number_page
            idx_par = index_paragraph
            break
        else:
            continue

for idx in range(idx_par, len(doc.Paragraphs)):
    range_par = doc.Paragraphs[idx].Range
    current_page_number = range_par.Information(WC.wdActiveEndPageNumber)

    text_par = range_par.Text.strip()

    # нашли заголовок
    if isHeading(range_par):
        par_format = range_par.ParagraphFormat

        if prevPage != current_page_number:
            prevPage = current_page_number

            processStep(pf.Alignment, WC.wdAlignParagraphCenter, pf.Alignment == WC.wdAlignParagraphCenter,
                        f"Выравнивание заголовка {i}")

            processStep(par_format.SpaceBefore, SPACE_BEFORE_HEADERS,
                        pc(par_format.SpaceBefore, SPACE_BEFORE_HEADERS, 0.3), f"Отступ снизу после заголовка {idx}")
        elif doc.Paragraphs[idx - 1].Range.Information(WC.wdActiveEndPageNumber) == current_page_number:
            print(Error(f"{current_page_number} стр. Заголовок не в начале страницы", text_par))
        else:
            print(Error(f"{current_page_number} стр. На этой странице уже есть заголовок", text_par))

for res in results:
    print(res)
