from pathlib import Path
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from PIL import Image
import subprocess
import tempfile
import shutil


# ============================================================
# 1. PATH CONFIGURATION
# ============================================================

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

# Slide indices are zero-based
IMAGE_SLIDE_INDEX = 1
IMAGE_SHAPE_ID = 10

VIDEO_SLIDE_INDEX = 4
VIDEO_SHAPE_ID = 6

# Thumbnail timestamp in seconds
THUMBNAIL_TIMESTAMP = 2


# ============================================================
# 2. VALIDATE INPUT FILES
# ============================================================

def validate_files():

    for path in [PPTX_PATH, IMAGE_PATH, VIDEO_PATH]:
        if not path.is_file():
            raise FileNotFoundError(
                f"File not found: {path}"
            )

    if shutil.which("ffmpeg") is None:
        raise RuntimeError(
            "FFmpeg not found. Install FFmpeg and add it to PATH."
        )

    print("All input files validated.")


# ============================================================
# 3. FIND SHAPE BY ID
# ============================================================

def find_shape(slide, shape_id):

    for shape in slide.shapes:
        if shape.shape_id == shape_id:
            return shape

    raise ValueError(
        f"Shape ID {shape_id} not found."
    )


# ============================================================
# 4. REPLACE ORIGINAL SHAPE
# ============================================================

def replace_shape(original_shape, new_shape):

    old_element = original_shape._element
    new_element = new_shape._element

    parent = old_element.getparent()
    index = parent.index(old_element)

    # Remove newly created element from its current location
    new_parent = new_element.getparent()
    new_parent.remove(new_element)

    # Insert at original stacking position
    parent.insert(index, new_element)

    # Remove original placeholder
    parent.remove(old_element)


# ============================================================
# 5. INSERT IMAGE
# ============================================================

def insert_image(slide, shape_id, image_path):

    shape = find_shape(slide, shape_id)

    print(f"\nImage placeholder found: {shape.name}")

    # Native picture placeholder
    if (
        shape.is_placeholder
        and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
    ):

        shape.insert_picture(str(image_path))

        print("Image inserted into picture placeholder.")
        return

    # Other placeholder types
    picture = slide.shapes.add_picture(
        str(image_path),
        shape.left,
        shape.top,
        width=shape.width,
        height=shape.height
    )

    replace_shape(shape, picture)

    print("Image inserted successfully.")


# ============================================================
# 6. EXTRACT VIDEO THUMBNAIL USING FFMPEG
# ============================================================

def extract_video_thumbnail(
    video_path,
    output_path,
    timestamp=2
):

    print(f"\nExtracting video thumbnail at {timestamp}s...")

    command = [
        "ffmpeg",
        "-y",
        "-ss", str(timestamp),
        "-i", str(video_path),
        "-frames:v", "1",
        "-q:v", "2",
        str(output_path)
    ]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"FFmpeg failed:\n{result.stderr}"
        )

    if not output_path.is_file():
        raise RuntimeError(
            "Thumbnail was not generated."
        )

    # Validate image
    with Image.open(output_path) as img:
        img.verify()

    print(f"Thumbnail generated: {output_path}")

    return output_path


# ============================================================
# 7. EMBED VIDEO WITH THUMBNAIL
# ============================================================

def insert_video(
    slide,
    shape_id,
    video_path,
    thumbnail_path
):

    shape = find_shape(slide, shape_id)

    print(f"\nVideo placeholder found: {shape.name}")

    print(
        f"Position: ({shape.left}, {shape.top})"
    )

    print(
        f"Size: ({shape.width}, {shape.height})"
    )

    movie = slide.shapes.add_movie(
        movie_file=str(video_path),
        left=shape.left,
        top=shape.top,
        width=shape.width,
        height=shape.height,
        poster_frame_image=str(thumbnail_path),
        mime_type="video/mp4"
    )

    replace_shape(shape, movie)

    print("Video embedded with thumbnail successfully.")


# ============================================================
# 8. MAIN AUTOMATION
# ============================================================

def main():

    print("=" * 60)
    print("CONTENT REPORT AUTOMATION")
    print("=" * 60)

    validate_files()

    # Open existing presentation
    prs = Presentation(str(PPTX_PATH))

    print(f"\nTotal slides: {len(prs.slides)}")

    # --------------------------------------------------------
    # INSERT IMAGE - SLIDE 2
    # --------------------------------------------------------

    image_slide = prs.slides[IMAGE_SLIDE_INDEX]

    insert_image(
        slide=image_slide,
        shape_id=IMAGE_SHAPE_ID,
        image_path=IMAGE_PATH
    )

    # --------------------------------------------------------
    # GENERATE THUMBNAIL AND INSERT VIDEO - SLIDE 4
    # --------------------------------------------------------

    video_slide = prs.slides[VIDEO_SLIDE_INDEX]

    # Temporary thumbnail, automatically cleaned up
    with tempfile.TemporaryDirectory() as temp_dir:

        thumbnail_path = (
            Path(temp_dir) / "video_thumbnail.jpg"
        )

        extract_video_thumbnail(
            video_path=VIDEO_PATH,
            output_path=thumbnail_path,
            timestamp=THUMBNAIL_TIMESTAMP
        )

        insert_video(
            slide=video_slide,
            shape_id=VIDEO_SHAPE_ID,
            video_path=VIDEO_PATH,
            thumbnail_path=thumbnail_path
        )

        # Save while temporary thumbnail still exists
        prs.save(str(OUTPUT_PATH))

    print("\n" + "=" * 60)
    print("AUTOMATION COMPLETED")
    print("=" * 60)

    print(f"Output file: {OUTPUT_PATH}")
    print(
        f"File size: "
        f"{OUTPUT_PATH.stat().st_size / (1024 * 1024):.2f} MB"
    )


# ============================================================
# 9. EXECUTE
# ============================================================

if __name__ == "__main__":
    main()
