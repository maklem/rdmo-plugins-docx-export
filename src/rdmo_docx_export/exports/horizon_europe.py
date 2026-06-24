from collections.abc import Callable
from importlib import resources
import io
import logging

from docx.enum.dml import MSO_COLOR_TYPE
from docx.text.run import Run
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph
from docx import Document

from django.http import HttpResponse
from django.utils import translation

from rdmo.projects.exports import Export
from rdmo.projects.models.value import Value
from rdmo.projects.managers import ValueQuerySet

import rdmo_docx_export.exports.templates as templates

logger = logging.getLogger("rdmo."+__name__)

class Style:
    def __init__(self, run: Run):
        self.style = run.style
        self.font = run.font.name
        self.color = run.font.color
        self.size = run.font.size
        self.shadow = run.font.shadow
        self.outline = run.font.outline

        try:
            self.highlight = run.font.highlight_color
        except ValueError:
            self.highlight = None
            
    def apply(self, run: Run) -> Run:
        run.style.base_style = self.style.base_style
        run.style.style_id = self.style.style_id

        run.font.name = self.font
        if self.color.type == MSO_COLOR_TYPE.RGB:
            run.font.color.rgb = self.color.rgb
        elif self.color.type == MSO_COLOR_TYPE.THEME:
            run.font.color.theme_color = self.color.theme_color
        
        if self.highlight is not None:
            run.font.highlight_color = self.highlight
        run.font.size = self.size
        run.font.shadow = self.shadow
        run.font.outline = self.outline

        return run



class _Context(object):
    def __init__(self, datasets: ValueQuerySet, funders: ValueQuerySet, partners: ValueQuerySet):
        self.datasets = datasets
        self.funders = funders
        self.partners = partners

_ParagraphFunction = Callable[['_Context', Paragraph],None]
_Replacements = dict[str, str|_ParagraphFunction|None]


def has_value(query_response: Value | None) -> bool:
    return query_response is not None and len(query_response.value.strip()) > 0


class HorizonEuropeDocxExport(Export):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def _stub(self, context: _Context, para: Paragraph) -> None:
        para.add_run("Stub Text. This is not implemented yet.")

    def _a1a(self, context: _Context, para: Paragraph) -> None:
        """
        Will you re-use any existing data and what will you re-use it for?
        """
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

    def _a1b(self, context: _Context, para: Paragraph) -> None:
        """
        State the reasons if re-use of any existing data has been considered but discarded.
        """
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

    def _a2(self, context: _Context, para: Paragraph) -> None:
        """
        What types and formats of data will the project generate or re-use?
        """
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

    def _a3(self, context: _Context, para: Paragraph) -> None:
        """
        What is the purpose of the data generation or re-use and its relation to the objectives of the project?
        """
        first = True
        for dataset in context.datasets:
            description = self.get_value("project/dataset/usage_description", set_index=dataset.set_index)

            if not description:
                continue

            if not first:
                para.add_run("\n\n")
            first = False

            headline = para.add_run(f"Dataset {dataset.value}:\n")
            headline.italic = True
            para.add_run(description.value)

    def _a4(self, context: _Context, para: Paragraph) -> None:
        """
        What is the expected size of the data that you intend to generate or re-use?
        """
        first = True
        for dataset in context.datasets:
            expected_size = self.get_value("project/dataset/size/volume", set_index=dataset.set_index)

            if not expected_size:
                continue

            if not first:
                para.add_run("\n\n")
            first = False

            headline = para.add_run(f"Dataset {dataset.value}: ")
            headline.italic = True
            para.add_run(f"The expected size of the data is {expected_size.value.lower()}.")

    def _a5(self, context: _Context, para: Paragraph) -> None:
        """
        What is the origin/provenance of the data, either generated or re-used?
        """
        first = True
        for dataset in context.datasets:
            provenance_content = self.get_values("project/dataset/provenance/content", set_index=dataset.set_index)
            origin = self.get_value("project/dataset/origin", set_index=dataset.set_index)

            if not provenance_content:
                continue

            if not first:
                para.add_run("\n\n")
            first = False

            headline = para.add_run(f"Dataset {dataset.value}: ")
            headline.italic = True
            headline.add_break()
            for v in provenance_content:
                para.add_run(v.text).add_break()

            if "reused" in origin.value.lower():
                author = self.get_value( 'project/dataset/creator/name', set_index=dataset.set_index)
                uri = self.get_value( 'project/dataset/uri', set_index=dataset.set_index)
                para.add_run(f"The data were created by {author.value} and can be found at the following address: {uri.value}")

    def _a6(self, context: _Context, para: Paragraph) -> None:
        """
        To whom might your data be useful ('data utility'), outside your project?
        """
        first = True
        for dataset in context.datasets:
            use_cases = self.get_values("project/dataset/reuse_scenario", set_index=dataset.set_index)

            if not use_cases:
                continue

            if not first:
                para.add_run("\n\n")
            first = False

            headline = para.add_run(f"Dataset {dataset.value}: ")
            headline.italic = True
            headline.add_break()
            for usecase in use_cases:
                para.add_run(" * " +usecase.text).add_break()

    def _a7(self, context: _Context, para: Paragraph) -> None:
        """
        Will the data be identified by a persistent identifier?
        """
        first = True
        for dataset in context.datasets:
            has_pid = self.get_bool("project/dataset/pids/yesno", set_index=dataset.set_index)
            pids = self.get_values("project/dataset/pids/system", set_index=dataset.set_index)

            if not has_pid:
                continue

            if not first:
                para.add_run("\n\n")
                first = False

            headline = para.add_run(f"Dataset {dataset.value}: ")
            headline.italic = True
            headline.add_break()
            for pid in pids:
                para.add_run(" * " +pid.value).add_break()

    def _a8(self, context: _Context, para: Paragraph) -> None:
        """
        Will rich metadata be provided to allow discovery? What metadata will be created? What disciplinary or general standards will be followed?
	    In case metadata standards do not exist in your discipline, please outline what type of metadata will be created and how.
        """
        first = True
        for dataset in context.datasets:
            automatic = self.get_value("project/dataset/metadata/creation_automatic", set_index=dataset.set_index)
            semiauto = self.get_value("project/dataset/metadata/creation_semi_automatic", set_index=dataset.set_index)
            manual = self.get_value("project/dataset/metadata/creation_manual", set_index=dataset.set_index)

            if not any([has_value(automatic), has_value(semiauto), has_value(manual)]):
                continue

            if not first:
                para.add_run("").add_break()
            first = False

            headline = para.add_run(f"Dataset {dataset.value}: ")
            headline.italic = True
            headline.add_break()
            if has_value(automatic):
                para.add_run(" * Automatically created: " + automatic.value).add_break()
            if has_value(semiauto):
                para.add_run(" * Automatically created, manually corrected: " + semiauto.value).add_break()
            if has_value(manual):
                para.add_run(" * Manually created: " + manual.value).add_break()

    def _a9(self, context: _Context, para: Paragraph) -> None:
        """
        Will search keywords be provided in the metadata to optimize the possibility for discovery and then potential re-use?
        """
        first = True
        for dataset in context.datasets:
            keywords = self.get_value("project/dataset/metadata/search_keywords", set_index=dataset.set_index)

            if not has_value(keywords):
                continue

            if not first:
                para.add_run("").add_break()
            first = False

            headline = para.add_run(f"Dataset {dataset.value}: ")
            headline.italic = True
            para.add_run(keywords.value)

    def _a10(self, context: _Context, para: Paragraph) -> None:
        """
        Will metadata be offered in such a way that they can be harvested and indexed?
        """
        first = True
        for dataset in context.datasets:
            harvesting = self.get_value("project/dataset/metadata/harvesting", set_index=dataset.set_index)

            if not has_value(harvesting):
                continue

            if not first:
                para.add_run("").add_break()
            first = False

            headline = para.add_run(f"Dataset {dataset.value}: ")
            headline.italic = True
            para.add_run(harvesting.value)

    def _a11(self, context: _Context, para: Paragraph) -> None:
        """
        Will the data be deposited in a trusted repository?
        """
        first = True
        for dataset in context.datasets:
            repositories = self.get_values("project/dataset/preservation/repository", set_index=dataset.set_index)
            trusted = self.get_values("project/dataset/preservation/trusted", set_index=dataset.set_index)

            if len(repositories) == 0:
                continue

            if not first:
                para.add_run("\n").add_break()
            first = False

            headline = para.add_run(f"Dataset {dataset.value}: ")
            headline.italic = True
            para.add_run("The dataset is stored in a ")
            para.add_run(", ".join(r.value for r in repositories) + ".")

            if len(trusted) > 0:
                para.add_run("\nThis repository is trusted because: ")
                para.add_run(", ".join(t.value for t in trusted) +  ".")

    def _a12(self, context: _Context, para: Paragraph) -> None:
        """
        Have you explored appropriate arrangements with the identified repository where your data will be deposited?
        """
        first = True
        for dataset in context.datasets:
            repository_arrangements = self.get_values("project/dataset/preservation/repository_arrangements", set_index=dataset.set_index)

            if len(repository_arrangements) == 0:
                continue

            if not first:
                para.add_run("\n").add_break()
            first = False

            headline = para.add_run(f"Dataset {dataset.value}: ")
            headline.italic = True
            para.add_run(", ".join(r.value for r in repository_arrangements) + ".")

    def _a13a(self, context: _Context, para: Paragraph) -> None:
        """
        Does the repository ensure that the data are assigned an identifier?
        """
        first = True
        for dataset in context.datasets:
            has_pid = self.get_value("project/dataset/pids/yesno", set_index=dataset.set_index)

            if not has_value(has_pid):
                continue

            if not first:
                para.add_run("\n").add_break()
            first = False

            headline = para.add_run(f"Dataset {dataset.value}: ")
            headline.italic = True
            para.add_run(has_pid.value)

    def _a13b(self, context: _Context, para: Paragraph) -> None:
        """
        Will the repository resolve the identifier to a digital object?
        """
        first = True
        for dataset in context.datasets:
            resolver = self.get_value("project/dataset/pids/resolver", set_index=dataset.set_index)

            if not has_value(resolver):
                continue

            if not first:
                para.add_run("\n").add_break()
            first = False

            headline = para.add_run(f"Dataset {dataset.value}: ")
            headline.italic = True
            para.add_run(resolver.value)
            if "no" not in resolver.value.lower():
                para.add_run(", the repository will resolve the identifier to a digital object")
            para.add_run(".")

    def _replace_paragraph_contents(self, replacements: _Replacements, context: _Context, para: Paragraph):
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
            if para.text in replacements:
                value = replacements[para.text]
                if value is None:
                    replace_para(para, "***error: value not defined***")
                elif isinstance(value, str):
                    replace_para(para,value)
                else:
                    style = Style(para.runs[0])
                    para.text = ""
                    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    value(context, para)
                    for run in para.runs:
                        style.apply(run)
            else:
                replace_para(para, "Lorem Ipsum... ***replacement sequence not defined***")

    def render(self):
        logger.info("Generating Docx Document...")
        logger.info("TEST logging Umlaute...")
        logger.info("ÄÖÜ...".encode('unicode_escape').decode('ascii'))
        logger.info(f"{'äöü'}")
        logger.info("TEST done.")

        template = resources.files(templates) / "horizon-template.docx"
        doc = Document(template.open("rb"))
        with translation.override("en"):
            context = _Context(
                datasets = self.get_set("project/dataset/id"),
                partners = self.get_set("project/partner/id"),
                funders = self.get_set("project/funder/id"),
            )

            replacements = {
                "{{projectnumber}}" : self.get_text("project/funder/grant_nr"),
                "{{projectacronym}}": self.get_text("project/acronym"),
                "{{projecttitle}}"  : self.get_text("project/title"),
                "{{dmpdate}}"       : self.get_text("project/dmp/dmp_date"),
                "{{dmpversion}}"    : self.get_text("project/dmp/dmp_version"),
                "{{Answer01a}}"     : self._a1a,
                "{{Answer01b}}"     : self._a1b,
                "{{Answer02}}"      : self._a2,
                "{{Answer03}}"      : self._a3,
                "{{Answer04}}"      : self._a4,
                "{{Answer05}}"      : self._a5,
                "{{Answer06}}"      : self._a6,
                "{{Answer07}}"      : self._a7,
                "{{Answer08}}"      : self._a8,
                "{{Answer09}}"      : self._a9,
                "{{Answer10}}"      : self._a10,
                "{{Answer11}}"      : self._a11,
                "{{Answer12}}"      : self._a12,
                "{{Answer13a}}"     : self._a13a,
                "{{Answer13b}}"     : self._a13b,
                "{{Answer14a}}"     : self._stub,
                "{{Answer14b}}"     : self._stub,
                "{{Answer15}}"      : self._stub,
                "{{Answer16}}"      : self._stub,
                "{{Answer17}}"      : self._stub,
                "{{Answer18}}"      : self._stub,
                "{{Answer19}}"      : self._stub,
                "{{Answer20}}"      : self._stub,
                "{{Answer21}}"      : self._stub,
                "{{Answer22}}"      : self._stub,
                "{{Answer23}}"      : self._stub,
                "{{Answer24}}"      : self._stub,
                "{{Answer25}}"      : self._stub,
                "{{Answer26}}"      : self._stub,
                "{{Answer27}}"      : self._stub,
                "{{Answer28}}"      : self._stub,
                "{{Answer29}}"      : self._stub,
                "{{Answer30}}"      : self._stub,
                "{{Answer31a}}"     : self._stub,
                "{{Answer31b}}"     : self._stub,
                "{{Answer31c}}"     : self._stub,
                "{{Answer32}}"      : self._stub,
                "{{Answer33}}"      : self._stub,
                "{{Answer34}}"      : self._stub,
                "{{Answer35}}"      : self._stub,
                "{{Answer36}}"      : self._stub,
                "{{Answer37}}"      : self._stub,
                "{{Answer38a}}"     : self._stub,
                "{{Answer38b}}"     : self._stub,
                "{{Answer39}}"      : self._stub,
                "{{Answer40}}"      : self._stub,
                "{{Answer41}}"      : self._stub,
                "{{Answer42}}"      : self._stub,
            }

            for para in doc.paragraphs:
                self._replace_paragraph_contents(replacements, context, para)
            for tab in doc.tables:
                for c in tab.columns:
                    for cell in c.cells:
                        for para in cell.paragraphs:
                            self._replace_paragraph_contents(replacements, context, para)

        response_data = io.BytesIO()
        doc.save(response_data)
        response_data.seek(0)  # Rewind file pointer to extract generated data!

        response = HttpResponse(response_data.read(), content_type='application/docx')
        response['Content-Disposition'] = 'attachment; filename="horizon-europe-export.docx"'
        return response

    def submit(self):
        raise NotImplementedError
