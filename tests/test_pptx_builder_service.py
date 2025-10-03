from pathlib import Path

from app.services.pptx_builder_service import PPTXBuilderService


class _DummyTextRange:
    def __init__(self, text: str) -> None:
        self.Text = text


class _DummyTextFrame:
    def __init__(self, text: str) -> None:
        self.TextRange = _DummyTextRange(text)


class _DummyShape:
    def __init__(self, text: str) -> None:
        self.TextFrame = _DummyTextFrame(text)
        self.Name = "DummyShape"


def test_detail_context_includes_metadata() -> None:
    service = PPTXBuilderService()

    detail = {
        "slide_number": 5,
        "name": "Alpha",
        "category": "Exploration",
    }

    context = service._detail_context(detail, "base slide")

    assert "slide #5" in context
    assert '"Alpha"' in context
    assert 'category "Exploration"' in context


def test_replace_placeholder_text_warns_when_missing(monkeypatch) -> None:
    service = PPTXBuilderService()
    dummy_shape = _DummyShape("Other text")

    monkeypatch.setattr(service, "_iter_text_shapes", lambda _slide: [dummy_shape])

    warnings: list[str] = []

    replaced = service._replace_placeholder_text(
        slide=None,
        placeholder="Name Candidate",
        replacement="Alpha",
        warnings=warnings,
        required=True,
        context="test slide",
    )

    assert replaced == 0


    def test_copy_slide_from_template_falls_back_to_insert(tmp_path: Path) -> None:
        template_path = tmp_path / "template.pptx"
        template_path.write_text("dummy")

        class _DummyTemplateSlide:
            def Copy(self):
                pass

        class _DummyTemplateSlides:
            def __call__(self, index: int):
                assert index == 1
                return _DummyTemplateSlide()

        class _DummyRange:
            def Item(self, idx: int):
                return f"slide-{idx}"

        class _DummySlides:
            def __init__(self) -> None:
                self.Count = 0
                self.insert_called = False

            def Paste(self):
                raise RuntimeError("Clipboard empty")

            def InsertFromFile(self, filename: str, index: int, slide_start: int = 1, slide_end: int = 1):
                self.insert_called = True
                assert filename == str(template_path.resolve())
                assert index == 0
                assert slide_start == 1 and slide_end == 1
                return _DummyRange()

        class _DummyPresentation:
            def __init__(self, slides) -> None:
                self.Slides = slides

        template_presentation = _DummyPresentation(_DummyTemplateSlides())
        target_slides = _DummySlides()
        target_presentation = _DummyPresentation(target_slides)

        slide = PPTXBuilderService._copy_slide_from_template(
            template_presentation,
            target_presentation,
            template_path,
        )

        assert slide == "slide-1"
        assert target_slides.insert_called is True
    assert warnings == ["Placeholder 'Name Candidate' not found in test slide."]
