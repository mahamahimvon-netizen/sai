# sai_srikar_Visualstroyboard_ModernBoho — generator

This repository contains a Python script to generate an editable 5-slide PowerPoint (PPTX) named exactly `sai_srikar_Visualstroyboard_ModernBoho.pptx` for the "Male Feminine Men — Modern Boho" moodboard.

What the script does
- Downloads free images using Unsplash Source queries
- Builds a 5-slide PPTX with:
  1) Title slide
  2) Mood & complementary colour palette
  3) Fabrics, textures & prints
  4) Outfits & silhouette inspiration
  5) Customer & story + image credits

How to run locally
1. Install Python 3.8+.
2. Clone this repo or download the files.
3. Install dependencies:
   ```
   pip install python-pptx Pillow requests
   ```
4. Run the generator:
   ```
   python generate_ppt.py
   ```
5. The script will create a `ppt_images` folder and save `sai_srikar_Visualstroyboard_ModernBoho.pptx` in the repo root.

Notes
- The script uses Unsplash Source (no API key) so images will vary each run.
- If you prefer to use specific images or photographer credits, replace the queries or manually add images to `ppt_images` and update the script.

Next steps
- After running the script, upload the PPTX to Google Drive and create a share link if you want it available elsewhere.
