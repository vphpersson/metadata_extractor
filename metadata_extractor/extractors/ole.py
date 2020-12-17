from pathlib import Path
from typing import Dict, Any

from olefile import OleFileIO

from metadata_extractor.extractors import MetadataExtractor, register_metadata_extractor


@register_metadata_extractor
class OLEMetadataExtractor(MetadataExtractor):

    MIME_TYPES = {
        'application/msword',
        'application/vnd.ms-excel',
        'application/vnd.ms-powerpoint',
    }

    @staticmethod
    def from_path(path: Path) -> Dict[str, Any]:
        with OleFileIO(path) as ole_file_io:
            return ole_file_io.get_metadata().__dict__.items()
