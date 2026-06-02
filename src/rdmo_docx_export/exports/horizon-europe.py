import io

from importlib import resources
from django.http import HttpResponse

from docx import Document

from rdmo.projects.exports import Export

import rdmo_docx_export.exports.templates as templates


class HorizonEuropeDocxExport(Export):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def render(self):
        template = resources.files(templates) / "horizon-template.docx"
        doc = Document(template.open("rb"))
        
        response_data = io.BytesIO()
        doc.save(response_data)

        response = HttpResponse(response_data, content_type='application/docx')
        return response

    def submit(self):
        raise NotImplementedError
