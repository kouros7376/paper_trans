import docx2txt
import PyPDF2
import json

def extract_docx(file_path):
    text = docx2txt.process(file_path)
    with open(file_path + ".txt", "w", encoding="utf-8") as f:
        f.write(text)
    print(f"Extracted {file_path}")

def extract_pdf(file_path):
    with open(file_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        text = ""
        for i in range(len(reader.pages)):
            text += reader.pages[i].extract_text()
    with open(file_path + ".txt", "w", encoding="utf-8") as f:
        f.write(text)
    print(f"Extracted {file_path}")

extract_docx("다우오피스_전자결재_연동_기술가이드.docx")
extract_pdf("전자결재 기안 – 다우오피스 차세대.pdf")
extract_docx("(양식)업무인수인계서 (1).docx")
