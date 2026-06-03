import io
from importlib import resources

from docx.enum.text import WD_BREAK, WD_ALIGN_PARAGRAPH
from docx.styles.style import CharacterStyle
from docx.text.paragraph import Paragraph
from docx import Document
from docx.table import Table

from django.http import HttpResponse
from django.utils import translation

from rdmo.projects.exports import Export

import rdmo_docx_export.exports.templates as templates


class HorizonEuropeDocxExport(Export):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def _a1a(self, datasets, para: Paragraph):
        para.text = ""
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        first = True
        for data in datasets:
            if not first:
                para.add_run("\n\n")
            first = False

            origin = self.get_value("project/dataset/origin", set_index=data.set_index)

            headline = para.add_run(f"Dataset {data.value}")
            headline.italic = True
            headline.add_break()
            para.add_run(f"This dataset is {origin.value.lower()}.").add_break()
            para.add_run(self.get_text("project/dataset/usage_description", set_index=data.set_index))

    def render(self):

        template = resources.files(templates) / "horizon-template.docx"
        doc = Document(template.open("rb"))
        with translation.override("en"):
            datasets = self.get_set("project/dataset/id")
            partners = self.get_set("project/partner/id")
            funders = self.get_set("project/funder/id")

            replacements = {
                "{{projectnumber}}": self.get_text("project/funder/grant_nr"),
                "{{projectacronym}}": self.get_text("project/acronym"),
                "{{projecttitle}}": self.get_text("project/title"),
                "{{dmpdate}}": self.get_text("project/dmp/dmp_date"),
                "{{dmpversion}}": self.get_text("project/dmp/dmp_version"),
                "{{Answer1a}}": self._a1a,
            }

            for para in doc.paragraphs:
                if para.text.startswith("{{") and para.text.endswith("}}"):
                    if para.text in replacements:
                        if isinstance(replacements[para.text], str):
                            para.text = replacements[para.text]
                        else:
                            replacements[para.text](datasets, para)
                    else:
                        para.text = "Lorem Ipsum..."

            for tab in doc.tables:
                for c in tab.columns:
                    for cell in c.cells:
                        if cell.text.startswith("{{") and cell.text.endswith("}}"):
                            print(f"Found template section {cell.text} in table")
                            if cell.text in replacements and replacements[cell.text] is not None:
                                if isinstance(replacements[cell.text], str):
                                    cell.text = replacements[cell.text]
                                else:
                                    replacements[cell.text](datasets, cell)
                            else:
                                cell.text = "Lorem Ipsum..."

        response_data = io.BytesIO()
        doc.save(response_data)
        response_data.seek(0)  # Rewind file pointer to extract generated data!

        response = HttpResponse(response_data.read(), content_type='application/docx')
        response['Content-Disposition'] = 'attachment; filename="horizon-europe-export.docx"'
        return response

    def submit(self):
        raise NotImplementedError
