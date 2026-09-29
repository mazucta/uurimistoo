"""Выгрузка материалов в файлы: презентация в .pptx, вопросы и идеи в .docx.

Текст берётся из того же JSON, который учитель правит на странице, поэтому файл всегда
совпадает с тем, что он видел. Оформление простое и читаемое: шрифт, размеры, отступы.
"""
import io

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor
from pptx import Presentation
from pptx.util import Cm as PCm, Pt as PPt

FONT = "Times New Roman"
SLIDE_FONT = "Arial"


def slides_pptx(title, slides, labels):
    """Презентация: титул, потом по слайду на пункт. Заметки учителя уходят в notes."""
    deck = Presentation()
    deck.slide_width, deck.slide_height = PCm(33.87), PCm(19.05)   # 16:9

    cover = deck.slides.add_slide(deck.slide_layouts[0])
    cover.shapes.title.text = title
    if len(cover.placeholders) > 1:
        cover.placeholders[1].text = labels.get("subtitle", "")

    for item in slides:
        slide = deck.slides.add_slide(deck.slide_layouts[1])
        slide.shapes.title.text = item.get("title", "")
        body = slide.placeholders[1].text_frame
        body.clear()
        points = [p for p in item.get("points", []) if p.strip()]
        for i, point in enumerate(points):
            para = body.paragraphs[0] if i == 0 else body.add_paragraph()
            para.text = point
            para.font.size = PPt(20)
            para.font.name = SLIDE_FONT
        if item.get("question"):
            para = body.add_paragraph()
            para.text = f"{labels['question']}: {item['question']}"
            para.font.size = PPt(18)
            para.font.italic = True
            para.font.name = SLIDE_FONT
        if item.get("notes"):
            slide.notes_slide.notes_text_frame.text = item["notes"]

    out = io.BytesIO()
    deck.save(out)
    return out.getvalue()


def start_doc(title):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(12)
    for section in doc.sections:
        section.left_margin = section.right_margin = Cm(2)
        section.top_margin = section.bottom_margin = Cm(2)
    head = doc.add_paragraph()
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = head.add_run(title)
    run.bold = True
    run.font.size = Pt(16)
    return doc


def quiz_docx(title, questions, labels):
    """Вопросы для класса, ответы для учителя отдельным разделом в конце."""
    doc = start_doc(title)
    letters = "ABCDEFGH"
    for i, q in enumerate(questions, 1):
        para = doc.add_paragraph()
        para.add_run(f"{i}. ").bold = True
        para.add_run(q.get("question", ""))
        mark = para.add_run(f"  [{labels.get(q.get('level'), '')}]")
        mark.font.size = Pt(9)
        mark.font.color.rgb = RGBColor(0x77, 0x77, 0x77)
        for n, option in enumerate(q.get("options", [])):
            line = doc.add_paragraph(f"{letters[n % len(letters)]}) {option}")
            line.paragraph_format.left_indent = Cm(1)
        if not q.get("options"):
            doc.add_paragraph("_" * 60).paragraph_format.left_indent = Cm(1)

    doc.add_page_break()
    head = doc.add_paragraph()
    head.add_run(labels["answers"]).bold = True
    for i, q in enumerate(questions, 1):
        doc.add_paragraph(f"{i}. {q.get('answer', '')}")
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def ideas_docx(title, ideas, labels):
    """Идеи заданий: по каждой шаги, что нужно, время и как проверить."""
    doc = start_doc(title)
    for i, idea in enumerate(ideas, 1):
        para = doc.add_paragraph()
        run = para.add_run(f"{i}. {idea.get('title', '')}")
        run.bold = True
        run.font.size = Pt(14)
        minutes = idea.get("minutes")
        if minutes:
            note = para.add_run(f"   {minutes} {labels['minutes']}")
            note.font.size = Pt(10)
            note.font.color.rgb = RGBColor(0x77, 0x77, 0x77)
        doc.add_paragraph(idea.get("what", ""))
        for key, name in (("needs", labels["needs"]), ("assess", labels["assess"])):
            if idea.get(key):
                line = doc.add_paragraph()
                line.add_run(f"{name}: ").bold = True
                line.add_run(idea[key])
        doc.add_paragraph()
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def build(kind, title, data, labels):
    """Файл материала: имя файла, тип и содержимое."""
    if kind == "slides":
        return "pptx", ("application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        slides_pptx(title, data.get("slides", []), labels))
    if kind == "quiz":
        return "docx", ("application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        quiz_docx(title, data.get("questions", []), labels))
    return "docx", ("application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    ideas_docx(title, data.get("ideas", []), labels))


if __name__ == "__main__":
    labels = {"question": "Вопрос классу", "subtitle": "Урок", "answers": "Ответы для учителя",
              "minutes": "мин", "needs": "Нужно", "assess": "Как проверить",
              "easy": "лёгкий", "medium": "средний", "hard": "трудный"}
    deck = slides_pptx("Тема урока", [{"title": "Первый слайд", "points": ["раз", "два"],
                                       "question": "почему?", "notes": "сказать вслух"}], labels)
    assert deck[:2] == b"PK" and len(deck) > 20000, len(deck)
    quiz = quiz_docx("Вопросы", [{"question": "Что это?", "kind": "choice", "options": ["а", "б"],
                                  "answer": "а", "level": "easy"}], labels)
    ideas = ideas_docx("Идеи", [{"title": "Аукцион", "what": "делают то-то", "needs": "доска",
                                 "minutes": 20, "assess": "сдают таблицу"}], labels)
    for blob in (quiz, ideas):
        assert blob[:2] == b"PK" and len(blob) > 10000, len(blob)
    kind, (mime, blob) = build("slides", "Тема", {"slides": []}, labels)
    assert kind == "pptx" and "presentation" in mime and blob[:2] == b"PK"
    print("ok")
