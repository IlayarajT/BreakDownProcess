import os
import zipfile


_LARGE_IMAGE_BYTES = 500 * 1024  # 500 KB


class DocxImageCleaner:
    """Scans a docx for oversized embedded images and reports them."""

    def __init__(self, filepath):
        self.filepath = filepath

    def process(self):
        """Return a list of image entry names that exceed the size threshold."""
        large = []
        try:
            with zipfile.ZipFile(self.filepath, "r") as z:
                for info in z.infolist():
                    if (
                        info.filename.startswith("word/media/")
                        and info.file_size > _LARGE_IMAGE_BYTES
                    ):
                        large.append(info.filename)
                        print(
                            f"[INFO] Large image: {os.path.basename(info.filename)} "
                            f"({info.file_size // 1024} KB)"
                        )
        except Exception as e:
            print(f"[WARN] DocxImageCleaner scan failed: {e}")
        return large
