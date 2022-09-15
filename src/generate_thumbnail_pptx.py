import collections  # ! supports pptx
import collections.abc  # ! supports pptx
import textwrap
from pathlib import PurePath

import pandas as pd
import pptx


def main():
    # ORGANIZATION CONFIGURATION
    ORGANIZATION_NAME = "UAB IT Research Computing"

    # FILE PATHS
    INPUT_CSV_NAME = PurePath("res") / "thumbnails.csv"
    INPUT_PPTX_TEMPLATE_NAME = PurePath("res") / "youtube-thumbnail-template.pptx"
    OUTPUT_PPTX_NAME = PurePath("thumbnails.pptx")

    # PRESENTATION CONFIGURATION
    PPTX_16_9_EMU = (12192000, 6868000)

    # LAYOUT
    TITLE_PLACEHOLDER_INDEX = 0
    TITLE_FONT_SIZE = pptx.util.Pt(48)  # type: ignore
    SUBTITLE_PLACEHOLDER_INDEX = 1
    SUBTITLE_FONT_SIZE = pptx.util.Pt(24)  # type: ignore

    # FIELDS
    TITLE = "title"
    PART = "part"
    CATEGORY = "category"
    INDEX = "index"
    DATE = "date"
    LAYOUT_INDEX = "layout_index"

    out = pptx.Presentation(INPUT_PPTX_TEMPLATE_NAME)

    out.slide_width = PPTX_16_9_EMU[0]
    out.slide_height = PPTX_16_9_EMU[1]

    df = pd.read_csv(INPUT_CSV_NAME)

    df[PART] = pd.to_numeric(df[PART])
    df[INDEX] = pd.to_numeric(df[INDEX])
    df[DATE] = pd.to_datetime(df[DATE])

    for _, row in df.iterrows():
        layout_index = row[LAYOUT_INDEX]
        slide_layout = out.slide_layouts[layout_index]
        slide = out.slides.add_slide(slide_layout)

        # TITLE
        part = row[PART]
        if not pd.isna(part):
            part = int(part)
            part_content = f" (Part {part:d})"
        else:
            part_content = ""

        title = row[TITLE]

        title_content = textwrap.dedent(
            f"""
            {title}{part_content}
            """
        ).strip()
        title_placeholder = slide.placeholders[TITLE_PLACEHOLDER_INDEX]
        title_placeholder.text = title_content
        for paragraph in title_placeholder.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = TITLE_FONT_SIZE

        # SUBTITLE
        category = row[CATEGORY]
        if not pd.isna(category):
            category_content = category
        else:
            category_content = ""

        index = row[INDEX]
        if not pd.isna(index):
            index = int(index)
            index_content = f" #{index:d}"
        else:
            index_content = ""

        date = row[DATE].strftime(r"%Y-%m-%d")

        subtitle_content = textwrap.dedent(
            f"""
            {category_content}{index_content}
            {ORGANIZATION_NAME}
            {date}
            """
        ).strip()
        subtitle_placeholder = slide.placeholders[SUBTITLE_PLACEHOLDER_INDEX]
        subtitle_placeholder.text = subtitle_content
        for paragraph in subtitle_placeholder.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = SUBTITLE_FONT_SIZE

    out.save(OUTPUT_PPTX_NAME)


if __name__ == "__main__":
    main()
