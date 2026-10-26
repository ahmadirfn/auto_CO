from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pathlib import Path

# ==================================================
# 1. PATH CONFIGURATION
# ==================================================

BASE_DIR = Path.cwd()

INPUT_DIR = BASE_DIR / "input"
ASSETS_DIR = INPUT_DIR / "assets"

PPTX_PATH = INPUT_DIR / "master_template.pptx"
OUTPUT_PATH = BASE_DIR / "output.pptx"

IMAGE_PATH = (
    ASSETS_DIR
    / "KVs"
    / "01 Front AMG Line"
    / "12344158_C.P.BH.X540_PAD_BEV_ProdStillUnreal_01-EXT-3-4-Front-AMG-Line_1-1_plate-DE_RGB preview.jpg"
)

VIDEO_PATH = (
    ASSETS_DIR
    / "Film_IAA"
    / "12344422_C.P.BH.X540_PAD_BEV_ProdFilm_07-INT-IAA-Trimpart-10s_1-1_Clean-NOSound_ProRes.mp4"
)

# Validate paths
for path in [PPTX_PATH, IMAGE_PATH, VIDEO_PATH]:
    if not path.is_file():
        raise FileNotFoundError(f"File not found: {path}")

# Open existing PowerPoint
prs = Presentation(PPTX_PATH)


# ==================================================
# 2. HELPER: FIND SHAPE BY ID
# ==================================================

def find_shape(slide, shape_id):
    for shape in slide.shapes:
        if shape.shape_id == shape_id:
            return shape

    raise ValueError(f"Shape ID {shape_id} not found")


# ==================================================
# 3. HELPER: REPLACE SHAPE WITH IMAGE
# ==================================================

def insert_image(slide, shape_id, image_path):

    shape = find_shape(slide, shape_id)

    print(f"Found image placeholder: {shape.name}")

    # Use native picture placeholder if supported
    if (
        shape.is_placeholder
        and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
    ):
        shape.insert_picture(str(image_path))
        print("Image inserted into picture placeholder.")
        return

    # Otherwise replace the existing shape
    left = shape.left
    top = shape.top
    width = shape.width
    height = shape.height

    old_element = shape._element
    parent = old_element.getparent()
    index = parent.index(old_element)

    picture = slide.shapes.add_picture(
        str(image_path),
        left,
        top,
        width=width,
        height=height
    )

    # Preserve stacking position
    parent.remove(picture._element)
    parent.insert(index, picture._element)

    # Remove original placeholder
    parent.remove(old_element)

    print("Image replaced existing placeholder.")


# ==================================================
# 4. HELPER: REPLACE SHAPE WITH VIDEO
# ==================================================

def insert_video(slide, shape_id, video_path):

    shape = find_shape(slide, shape_id)

    print(f"Found video placeholder: {shape.name}")

    left = shape.left
    top = shape.top
    width = shape.width
    height = shape.height

    old_element = shape._element
    parent = old_element.getparent()
    index = parent.index(old_element)

    movie = slide.shapes.add_movie(
        movie_file=str(video_path),
        left=left,
        top=top,
        width=width,
        height=height,
        mime_type="video/mp4"
    )

    # Preserve stacking position
    parent.remove(movie._element)
    parent.insert(index, movie._element)

    # Remove original placeholder
    parent.remove(old_element)

    print("Video embedded successfully.")


# ==================================================
# 5. INSERT ASSETS
# ==================================================

# Slide 2 - Image
insert_image(
    slide=prs.slides[1],
    shape_id=10,
    image_path=IMAGE_PATH
)

# Slide 4 - Video
insert_video(
    slide=prs.slides[4],
    shape_id=6,
    video_path=VIDEO_PATH
)


# ==================================================
# 6. SAVE PRESENTATION
# ==================================================

prs.save(str(OUTPUT_PATH))

print(f"\nPowerPoint saved successfully: {OUTPUT_PATH}")