import pytest

from donation_tracker.pdf_render import EdgeRenderer, find_edge


def test_renderer_without_edge_returns_false(tmp_path):
    renderer = EdgeRenderer(exe="")
    html = tmp_path / "a.html"
    html.write_text("<p>x</p>", encoding="utf-8")
    assert renderer.render(html, tmp_path / "a.pdf") is False


@pytest.mark.skipif(find_edge() is None, reason="Microsoft Edge not installed")
def test_edge_renders_real_pdf(tmp_path):
    html = tmp_path / "email.html"
    html.write_text("<html><body><h1>Receipt $50.00</h1></body></html>", encoding="utf-8")
    pdf = tmp_path / "out" / "email.pdf"
    assert EdgeRenderer().render(html, pdf) is True
    assert pdf.read_bytes().startswith(b"%PDF")
