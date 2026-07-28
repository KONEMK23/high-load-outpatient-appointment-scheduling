from __future__ import annotations

import os
from pathlib import Path

from pypdf import PdfReader, PdfWriter


def project_root() -> Path:
    configured = os.environ.get("MATH6185_PROJECT_ROOT")
    if configured:
        return Path(configured).resolve()
    candidate = Path(__file__).resolve().parents[2]
    if (candidate / "outputs").exists():
        return candidate
    return candidate.parents[1]


def paths() -> tuple[Path, Path]:
    output_dir = project_root() / "outputs" / "zw1f25_MATH6185_presentation"
    return (
        output_dir / "zw1f25_MATH6185_Presentation.pdf",
        output_dir / "zw1f25_MATH6185_Presentation_8_Slides.pdf",
    )


def build_pdf(source: Path, target: Path) -> None:
    reader = PdfReader(str(source))
    if len(reader.pages) != 14:
        raise ValueError(f"Expected 14 source pages, found {len(reader.pages)}")
    writer = PdfWriter()
    for page in reader.pages[:8]:
        writer.add_page(page)
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("wb") as stream:
        writer.write(stream)

    output = PdfReader(str(target))
    if len(output.pages) != 8:
        raise ValueError(f"Expected 8 output pages, found {len(output.pages)}")
    for index, (source_page, output_page) in enumerate(
        zip(reader.pages[:8], output.pages), start=1
    ):
        if source_page.mediabox != output_page.mediabox:
            raise ValueError(f"Page {index} media box changed")
        if source_page.cropbox != output_page.cropbox:
            raise ValueError(f"Page {index} crop box changed")


def main() -> None:
    source, target = paths()
    build_pdf(source, target)
    print(target)


if __name__ == "__main__":
    main()
