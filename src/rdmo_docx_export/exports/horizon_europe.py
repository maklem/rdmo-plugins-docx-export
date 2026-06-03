import io

from importlib import resources
from django.http import HttpResponse

from docx import Document
from docx.table import Table

from rdmo.projects.exports import Export

import rdmo_docx_export.exports.templates as templates


class HorizonEuropeDocxExport(Export):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def render(self):
        template = resources.files(templates) / "horizon-template.docx"
        doc = Document(template.open("rb"))
        
        for para in doc.paragraphs:
            if para.text.startswith("{{") and para.text.endswith("}}"):
                # print(f"Found template section {para.text}")
                para.text = "Lorem Ipsum..."

        for tab in doc.tables:
            for c in tab.columns:
                for cell in c.cells:
                    if cell.text.startswith("{{") and cell.text.endswith("}}"):
                        print(f"Found template section {cell.text} in table")
                        cell.text = "Lorem Ipsum..."

        response_data = io.BytesIO()
        doc.save(response_data)
        response_data.seek(0)  # Rewind file pointer to extract generated data!

        response = HttpResponse(response_data.read(), content_type='application/docx')
        response['Content-Disposition'] = 'attachment; filename="horizon-europe-export.docx"'
        return response

    def submit(self):
        raise NotImplementedError
