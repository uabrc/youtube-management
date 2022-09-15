# Scripts to Facilitate Youtube Channel Management

Channel: <https://www.youtube.com/channel/UCZoOS2e699Ge0DND1oy1BJQ>

Use Anaconda to set up the environment in `env.yml`.

## generate_thumbnail_pptx.py

To use:

1. Modify `res/thumbnails.csv` to match the list of videos in the channel.
    1. The smoothest experience is had if the entries are in the reverse order they appear in YouTube Studio Content Manager.
    2. Try to keep the thumbnails of Data Science Journal Club, Researcher Training Sessions, and other Videos visually distinct.
        1. DSJC - Layout Index 1
        2. RTS - Layout Index 4
        3. Others - Layout Index 2
    3. Use the date of streaming or creation, if possible. Don't assume the upload date is correct.
2. Use `python generate_thumbnail_pptx.py` to generate `thumbnails.pptx`.
3. Using PowerPoint, open `thumbnails.pptx`.
4. Save slides as images.
    1. Click "Save As..." to open a new dialog box.
    2. In the "Save as type" dropdown, Select "PNG Portable Network Graphics Format (*.png)".
    3. Click "Save" to open a new dialog box.
    4. Click "All Slides".
    5. Wait for thumbnails to be saved. They will be saved to the folder `thumbnails`.
5. Navigate to the `thumbnails` folder to find the thumbnails.
6. For each video, add the appropriate thumbnail.
