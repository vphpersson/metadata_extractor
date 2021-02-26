from pathlib import Path
from typing import Union

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import PDFObjRef, resolve1

from metadata_extractor.extractors import MetadataExtractor, register_metadata_extractor


@register_metadata_extractor
class PDFMetadataExtractor(MetadataExtractor):
    MIME_TYPES = {'application/pdf'}

    @staticmethod
    def from_path(path: Path) -> dict[str, bytes]:
        with path.open('rb') as fp:
            info: dict[str, Union[PDFObjRef, bytes]] = next(iter(PDFDocument(PDFParser(fp)).info), dict())
            return {
                key: resolve1(value) if isinstance(value, PDFObjRef) else value
                for key, value in info.items()
            }
