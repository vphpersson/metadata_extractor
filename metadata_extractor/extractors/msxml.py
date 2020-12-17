from pathlib import Path
from zipfile import ZipFile, Path as ZipPath
from typing import Dict, Any
from html.parser import HTMLParser
from logging import getLogger

from metadata_extractor.extractors import MetadataExtractor, register_metadata_extractor

LOG = getLogger(__name__)


class _OfficeOpenXMLParser(HTMLParser):
    """
    A parser for _Office Open XML_ data.

    Metadata appears the be enclosed within tags. The last encounter tag is noted, so that when data is encountered
    it can be mapped to a certain field.
    """

    def error(self, message):
        LOG.error(msg=message)
        if self.exception_on_error:
            raise ValueError(message)

    def __init__(self, *args, exception_on_error: bool = True, **kwargs):
        super().__init__(*args, **kwargs)
        self._last_seen_tag = None
        self.exception_on_error = exception_on_error

        self.metadata = dict()

    def handle_starttag(self, tag, attrs):
        self._last_seen_tag = tag

    def handle_data(self, data):
        if self._last_seen_tag:
            self.metadata[self._last_seen_tag] = data

    def extract_metadata(self, data):
        self.feed(data)
        return self.metadata


@register_metadata_extractor
class MSXMLMetadataExtractor(MetadataExtractor):

    MIME_TYPES = {
        'application/vnd.ms-excel.addin.macroEnabled.12',
        'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
        'application/vnd.ms-excel.sheet.macroEnabled.12',
        'application/vnd.ms-excel.template.macroEnabled.12',
        'application/vnd.ms-powerpoint.addin.macroEnabled.12',
        'application/vnd.ms-powerpoint.presentation.macroEnabled.12',
        'application/vnd.ms-powerpoint.slideshow.macroEnabled.12',
        'application/vnd.ms-powerpoint.template.macroEnabled.12',
        'application/vnd.ms-word.document.macroEnabled.12',
        'application/vnd.ms-word.template.macroEnabled.12',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'application/vnd.openxmlformats-officedocument.presentationml.slideshow',
        'application/vnd.openxmlformats-officedocument.presentationml.template',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
    }

    @staticmethod
    def from_path(path: Path) -> Dict[str, Any]:
        with ZipFile(path) as zip_file:
            return {
                doc_props_file_path.name: _OfficeOpenXMLParser().extract_metadata(doc_props_file_path.read_text())
                for doc_props_file_path in ZipPath(zip_file, 'docProps/').iterdir()
            }
