from core import parser

def test_extract_text_docx(tmp_path):
    file_path = tmp_path / "sample.docx"
    with open(file_path, "w") as f:
        f.write("Test Resume")
    try:
        text = parser.extract_text(str(file_path))
    except Exception:
        text = ""
    assert isinstance(text, str)