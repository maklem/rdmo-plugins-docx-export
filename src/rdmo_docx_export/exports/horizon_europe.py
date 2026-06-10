from typing import Callable, Any
from docx.enum.dml import MSO_COLOR_TYPE
from docx.text.run import Run
import io
from importlib import resources

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph
from docx import Document

from django.http import HttpResponse
from django.utils import translation

from rdmo.projects.exports import Export

import rdmo_docx_export.exports.templates as templates


class Style:
    def __init__(self, run: Run):
        self.style = run.style
        self.font = run.font.name
        self.color = run.font.color
        self.size = run.font.size
        self.shadow = run.font.shadow
        self.highlight = run.font.highlight_color
        self.outline = run.font.outline

    def apply(self, run: Run) -> Run:
        run.style.base_style = self.style.base_style
        run.style.style_id = self.style.style_id

        run.font.name = self.font
        if self.color.type == MSO_COLOR_TYPE.RGB:
            run.font.color.rgb = self.color.rgb
        elif self.color.type == MSO_COLOR_TYPE.THEME:
            run.font.color.theme_color = self.color.theme_color
        run.font.highlight_color = self.highlight
        run.font.size = self.size
        run.font.shadow = self.shadow
        run.font.outline = self.outline

        return run


_Replacements = dict[str, str|Callable[['_Context', Paragraph],None]|None]

class _Context(object):
    def __init__(self):
        self.datasets: Any = None
        self.funders: Any = None
        self.partners: Any = None
        self.replacements: _Replacements = {}

class HorizonEuropeDocxExport(Export):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def _a1a(self, context: _Context, para: Paragraph):
        para.text = ""
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        first = True
        for dataset in context.datasets:
            if not first:
                para.add_run("\n\n")
            first = False

            origin = self.get_value("project/dataset/origin", set_index=dataset.set_index)

            headline = para.add_run(f"Dataset {dataset.value}")
            headline.italic = True
            headline.add_break()
            para.add_run(f"This dataset is {origin.value.lower().replace('both (', '').replace(')','')}.\n")
            para.add_run(self.get_text("project/dataset/usage_description", set_index=dataset.set_index))

    def _a1b(self, context: _Context, para: Paragraph):
        para.text = ""
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        first = True
        for dataset in context.datasets:
            existing = self.get_value("project/dataset/reuse_existing", set_index=dataset.set_index)

            if existing and existing.value:  # if value exists and is not empty
                if not first:
                    para.add_run("\n\n")
                first = False

                headline = para.add_run(f"Dataset {dataset.value}:")
                headline.italic = True
                para.add_run(existing.value)

    def _a2(self, context: _Context, para: Paragraph):
        para.text = ""
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        first = True
        for dataset in context.datasets:
            description = self.get_value("project/dataset/description", set_index=dataset.set_index)
            format = self.get_value("project/dataset/format", set_index=dataset.set_index)

            if description or format:
                if not first:
                    para.add_run("\n\n")
                first = False

                headline = para.add_run(f"Dataset {dataset.value}:\n")
                headline.italic = True
            if description:
                para.add_run("The data are ")
                para.add_run(description.value).add_break()
            if format:
                para.add_run("They are provided in the following formats: ")
                para.add_run(format.value)

    def _replace_paragraph_contents(self, context: _Context, para: Paragraph):
        """
        Checks if a paragraph's content is to be replaced.

        If yes: Replaces contents. Tries to keep style intact.

        Functional Style elements (i.e. italic, bold) may be applied by content
        functions and will not be overwritten afterwards.
        Then some style elements are not yet copied correctly (Shadow, Outline, Theme).
        They are not readable/writable with python-docx, but we need to
        create 'runs' to apply functional style.
        """
        def replace_para(para, text):
            for run in para.runs:
                run.text = ""
            para.runs[0].text = text


        if para.text.startswith("{{") and para.text.endswith("}}"):
            if para.text in context.replacements:
                value = context.replacements[para.text]
                if value is None:
                    replace_para(para, "***error: value not defined***")
                elif isinstance(value, str):
                    replace_para(para,value)
                else:
                    style = Style(para.runs[0])
                    value(context, para)
                    for run in para.runs:
                        style.apply(run)
            else:
                replace_para(para, "Lorem Ipsum... ***replacement sequence not defined***")

    def render(self):
        context = _Context()
        template = resources.files(templates) / "horizon-template.docx"
        doc = Document(template.open("rb"))
        with translation.override("en"):
            context.datasets = self.get_set("project/dataset/id")
            context.partners = self.get_set("project/partner/id")
            context.funders = self.get_set("project/funder/id")

            context.replacements = {
                "{{projectnumber}}" : self.get_text("project/funder/grant_nr"),
                "{{projectacronym}}": self.get_text("project/acronym"),
                "{{projecttitle}}"  : self.get_text("project/title"),
                "{{dmpdate}}"       : self.get_text("project/dmp/dmp_date"),
                "{{dmpversion}}"    : self.get_text("project/dmp/dmp_version"),
                "{{Answer01a}}"     : self._a1a,
                "{{Answer01b}}"     : self._a1b,
                "{{Answer02}}"      : self._a2,
                "{{Answer03a}}"     : self._a3a,
                "{{Answer03b}}"     : self._a3b,
                "{{Answer03c}}"     : self._a3c,
                "{{Answer04}}"      : self._a2,
                "{{Answer05a}}"     : self._a1a,
                "{{Answer05b}}"     : self._a1b,
                "{{Answer06}}"      : self._a2,
                "{{Answer07}}"      : self._a2,
                "{{Answer08}}"      : self._a2,
                "{{Answer09}}"      : self._a2,
                "{{Answer10}}"      : self._a2,
                "{{Answer11}}"      : self._a2,
                "{{Answer12}}"      : self._a2,
                "{{Answer13}}"      : self._a2,
                "{{Answer14a}}"     : self._a3a,
                "{{Answer14b}}"     : self._a3b,
                "{{Answer14c}}"     : self._a3c,
                "{{Answer15}}"      : self._a2,
                "{{Answer16}}"      : self._a2,
                "{{Answer17}}"      : self._a2,
                "{{Answer18}}"      : self._a2,
                "{{Answer19}}"      : self._a2,
                "{{Answer20}}"      : self._a2,
                "{{Answer21}}"      : self._a2,
                "{{Answer22}}"      : self._a2,
                "{{Answer23}}"      : self._a2,
                "{{Answer24}}"      : self._a2,
                "{{Answer25}}"      : self._a2,
                "{{Answer26}}"      : self._a2,
                "{{Answer27}}"      : self._a2,
                "{{Answer28}}"      : self._a2,
                "{{Answer29}}"      : self._a2,
                "{{Answer30}}"      : self._a2,
                "{{Answer31a}}"     : self._a3a,
                "{{Answer31b}}"     : self._a3b,
                "{{Answer31c}}"     : self._a3c,
                "{{Answer32}}"      : self._a2,
                "{{Answer33}}"      : self._a2,
                "{{Answer34}}"      : self._a2,
                "{{Answer35}}"      : self._a2,
                "{{Answer36}}"      : self._a2,
                "{{Answer37}}"      : self._a2,
                "{{Answer38a}}"     : self._a1a,
                "{{Answer38b}}"     : self._a1b,
                "{{Answer39}}"      : self._a2,
                "{{Answer40}}"      : self._a2,
                "{{Answer41}}"      : self._a2,
                "{{Answer42}}"      : self._a2,
            }

            for para in doc.paragraphs:
                self._replace_paragraph_contents(context, para)
            for tab in doc.tables:
                for c in tab.columns:
                    for cell in c.cells:
                        for para in cell.paragraphs:
                            self._replace_paragraph_contents(context, para)

        response_data = io.BytesIO()
        doc.save(response_data)
        response_data.seek(0)  # Rewind file pointer to extract generated data!

        response = HttpResponse(response_data.read(), content_type='application/docx')
        response['Content-Disposition'] = 'attachment; filename="horizon-europe-export.docx"'
        return response

    def submit(self):
        raise NotImplementedError
