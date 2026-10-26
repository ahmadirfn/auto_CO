"""Generate the Mercedes-Benz content report from local media and two Excel sources.

Project layout:
  project/input/master_template.pptx
  project/input/assets_table1_cleaned.xlsx
  project/input/Export_report_20260917_113237_Raw(1).xlsx
  project/input/Assets/** (images and videos)
  project/automation-contentReport-main1/main_automation.py  (or project/main_automation.py)

Run: python main_automation.py
Plan without local media: python main_automation.py --plan-only
Dependencies: pip install python-pptx pandas openpyxl pillow imageio-ffmpeg
"""
from __future__ import annotations

import argparse
import csv
import math
import re
import subprocess
import tempfile
from collections import defaultdict
from pathlib import Path

import pandas as pd
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt

IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tif', '.tiff', '.webp'}
VIDEO_EXTS = {'.mp4', '.mov', '.m4v', '.avi', '.wmv', '.webm', '.mkv'}
IMAGE_SLOTS = {'9:16': 3, '16:9': 11, '4:3': 12, '4:5': 13}
TABLE_HEADERS = ['Asset ID', 'Color Space', 'Aspect Ratio', 'Format', 'Symphony Link']


def project_root():
    here = Path(__file__).resolve().parent
    for folder in (here, here.parent):
        if (folder / 'input' / 'master_template.pptx').is_file():
            return folder
    raise FileNotFoundError('Could not find input/master_template.pptx beside or above this script')

#extract/amik asset_id dari filename
def asset_id(filename):
    match = re.match(r'^(\d+)', str(filename).strip())
    return match.group(1) if match else ''


def parse_ratio(value):
    try:
        a, b = str(value).strip().split(':')
        return float(a) / float(b)
    except (ValueError, ZeroDivisionError):
        return None


def media_kind(row):
    fmt = str(row.get('format', '')).lower().strip().lstrip('.')
    ext = Path(str(row['file_name'])).suffix.lower()
    if fmt in ('image', 'picture') or '.' + fmt in IMAGE_EXTS or ext in IMAGE_EXTS:
        return 'image'
    if fmt == 'video' or '.' + fmt in VIDEO_EXTS or ext in VIDEO_EXTS:
        return 'video'
    return None


def read_assets(excel, raw):
    df = pd.read_excel(excel, dtype=str).fillna('')                                 #read  excel filename
    df.columns = df.columns.astype(str).str.strip().str.lower()
    required = {'file_name', 'format', 'aspect_ratio'}
    if not required.issubset(df.columns):
        raise ValueError(f'Missing columns: {sorted(required - set(df.columns))}')
    meta = pd.read_excel(raw, dtype={'Asset ID': str}).fillna('')                   #read raw excel filename & get metadata into memory
    meta.columns = meta.columns.astype(str).str.strip()
    if 'Asset ID' not in meta.columns:
        raise ValueError('Raw export requires Asset ID column')
    lookup = {str(r['Asset ID']).strip(): r for _, r in meta.iterrows()}            #build a lookup / merge between 2 excel files
    records = []
    for i, r in df.iterrows():
        if not str(r['file_name']).strip():
            continue
        kind = media_kind(r)
        if not kind:
            print(f'SKIP unknown format, Excel row {i+2}: {r["file_name"]}')
            continue
        aid = asset_id(r['file_name'])                                                  #find matching row within asset_id 
        info = lookup.get(aid, {})
        instrument = str(info.get('Instruments', '')).strip() or 'Unknown Instrument'
        records.append(dict(index=i+2, name=str(r['file_name']).strip(), id=aid,
                            kind=kind, ratio=str(r['aspect_ratio']).strip(),
                            instrument=instrument, raw=info, source=r))
    return records


def find_local_files(records, folder):
    files = sorted((p for p in folder.rglob('*') if p.is_file() and p.suffix.lower() in IMAGE_EXTS | VIDEO_EXTS), key=lambda p: str(p).lower())
    by_name = defaultdict(list)
    by_stem = defaultdict(list)
    by_id = defaultdict(list)
    for p in files:
        by_name[p.name.casefold()].append(p)
        by_stem[p.stem.casefold()].append(p)
        by_id[asset_id(p.name)].append(p)
    for r in records:
        n = r['name'].casefold()
        stem = Path(n).stem if Path(n).suffix.lower() in IMAGE_EXTS | VIDEO_EXTS else n
        candidates = by_name[n] or by_stem[stem]
        if not candidates:
            # The metadata export sometimes truncates long filenames.
            candidates = [p for p in by_id[r['id']] if p.name.casefold().startswith(stem)]
        candidates = [p for p in candidates if (p.suffix.lower() in IMAGE_EXTS) == (r['kind'] == 'image')]
        if len(candidates) == 1:
            r['path'] = candidates[0]
        else:
            r['path'] = None
            r['problem'] = 'missing' if not candidates else 'ambiguous: ' + '; '.join(str(p) for p in candidates[:5])
    return records


def group_and_allocate(records):
    # Group by full Instruments string. Within each group, reserve exact slots before fallbacks.
    groups = defaultdict(list)
    for r in records:
        if r.get('path'):
            groups[r['instrument']].append(r)
    plan = []
    for instrument, items in groups.items():
        non_square = [r for r in items if r['kind'] == 'image' and r['ratio'] != '1:1']
        squares = [r for r in items if r['kind'] == 'image' and r['ratio'] == '1:1']
        videos = [r for r in items if r['kind'] == 'video']
        image_pages = []
        # First allocate exact / near exact ratio slots, then unmatched assets.
        matched, fallback = [], []
        for r in non_square:
            value = parse_ratio(r['ratio'])
            exact = [s for s in IMAGE_SLOTS if value and abs(math.log(value / parse_ratio(s))) <= 0.035]
            (matched if exact else fallback).append((r, exact[0] if exact else None))
        for r, preferred in matched + fallback:
            candidates = [preferred] if preferred else sorted(IMAGE_SLOTS, key=lambda s: abs(math.log((parse_ratio(r['ratio']) or 1) / parse_ratio(s))))
            page = next(((p, s) for p in image_pages for s in candidates if s not in p), None)
            if page is None:
                if preferred:
                    slot = preferred
                else:
                    slot = candidates[0]
                new_page = {}
                image_pages.append(new_page)
                page = new_page, slot
            page[0][page[1]] = r
        for p in image_pages:
            plan.append(dict(layout=3, instrument=instrument, slots=p, assets=list(p.values())))
        for r in squares:
            plan.append(dict(layout=4, instrument=instrument, slots={'1:1': r}, assets=[r]))
        for r in videos:
            plan.append(dict(layout=5, instrument=instrument, slots={'video': r}, assets=[r]))
    # Preserve template section ordering: all slide-3 variants, then slide-4, then slide-5.
    return sorted(plan, key=lambda p: p['layout'])


def find_shape(slide, shape_id):
    for s in slide.shapes:
        if s.shape_id == shape_id:
            return s
    raise ValueError(f'Shape ID {shape_id} not found on slide')


def replace_shape(shape, new_shape):
    old = shape._element
    parent = old.getparent()
    index = parent.index(old)
    fresh = new_shape._element
    fresh.getparent().remove(fresh)
    parent.insert(index, fresh)
    parent.remove(old)


def contain_rect(shape, image_path):
    with Image.open(image_path) as im:
        w, h = im.size
    factor = min(shape.width / w, shape.height / h)
    width, height = int(w * factor), int(h * factor)
    return (shape.left + (shape.width-width)//2, shape.top + (shape.height-height)//2, width, height)


def put_image(slide, shape_id, path):
    sh = find_shaokaype(slide, shape_id)
    x, y, w, h = contain_rect(sh, path)
    pic = slide.shapes.add_picture(str(path), x, y, width=w, height=h)
    replace_shape(sh, pic)


def ffmpeg_path():
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except ImportError:
        import shutil
        executable = shutil.which('ffmpeg')
        if executable:
            return executable
        raise RuntimeError('Install imageio-ffmpeg: pip install imageio-ffmpeg')


def make_poster(video, destination):
    # Try the 2-second frame first, then the first frame for short videos.
    # A timeout prevents one problematic file from blocking the entire run.
    for timestamp in ('2', '0'):
        destination.unlink(missing_ok=True)
        command = [
            ffmpeg_path(), '-nostdin', '-hide_banner', '-loglevel', 'error',
            '-y', '-ss', timestamp, '-i', str(video),
            '-frames:v', '1', str(destination)
        ]
        try:
            result = subprocess.run(
                command, capture_output=True, text=True, timeout=60
            )
        except subprocess.TimeoutExpired:
            print(f'WARNING: Thumbnail timed out at {timestamp}s: {video.name}', flush=True)
            continue
        if result.returncode == 0 and destination.is_file() and destination.stat().st_size > 0:
            return
        print(f'WARNING: Thumbnail failed at {timestamp}s: {video.name} '
              f'{result.stderr[-250:]}', flush=True)
    raise RuntimeError(f'Unable to extract thumbnail from: {video}')


def put_video(slide, shape_id, path, temp):
    sh = find_shape(slide, shape_id)
    poster = temp / (asset_id(path.name) + '_poster.jpg')
    make_poster(path, poster)
    # Fit the video inside its placeholder, without stretching or cropping.
    with Image.open(poster) as im:
        width_px, height_px = im.size
    factor = min(sh.width / width_px, sh.height / height_px)
    width, height = int(width_px * factor), int(height_px * factor)
    left = sh.left + (sh.width - width) // 2
    top = sh.top + (sh.height - height) // 2
    movie = slide.shapes.add_movie(str(path), left, top, width, height,
                                   poster_frame_image=str(poster), mime_type='video/mp4')
    replace_shape(sh, movie)


def put_title(slide, instrument):
    sh = find_shape(slide, 2)
    if sh.has_text_frame:
        sh.text = instrument
        for paragraph in sh.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(17 if len(instrument) > 48 else 24)


def put_table(slide, shape_id, assets, layout):
    sh = find_shape(slide, shape_id)
    headers = TABLE_HEADERS if layout in (3, 4) else ['Asset ID', 'Variation', 'Quality', 'Symphony Link']
    data = [headers]
    for r in assets:
        info = r['raw']
        link = str(info.get('Hyperlink', '')).strip()                                                           #get the raw excel information
        if layout in (3, 4):
            data.append([r['id'], str(info.get('ColorSpace', '')).strip() or '[Color Space]',
                         r['ratio'], Path(r['path']).suffix.lstrip('.').upper(), link or '[Symphony Link]'])        #all of the metadata that is not in cleaned excel to put into pptx table
        else:
            data.append([r['id'], str(info.get('Perspective', '')).strip() or '[Variation]',
                         '[Quality]', link or '[Symphony Link]'])
    frame = slide.shapes.add_table(len(data), len(headers), sh.left, sh.top, sh.width, sh.height)
    replace_shape(sh, frame)
    table = frame.table
    for ri, row in enumerate(data):
        for ci, val in enumerate(row):
            cell = table.cell(ri, ci)
            cell.text = str(val)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Pt(3)
            cell.margin_top = cell.margin_bottom = Pt(2)
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER
                for run in paragraph.runs:
                    run.font.size = Pt(8 if layout == 3 else 9)
                    run.font.bold = ri == 0
                    if ri > 0 and ci == len(headers)-1 and str(val).startswith('https://'):
                        run.text = 'Symphony Link'
                        run.hyperlink.address = str(val)
    if layout in (3, 4):
        ratios = [0.17, 0.18, 0.19, 0.14, 0.32]
    else:
        ratios = [0.20, 0.35, 0.15, 0.30]
    widths = [int(sh.width * f) for f in ratios[:-1]]
    widths.append(sh.width-sum(widths))
    for col, width in zip(table.columns, widths):
        col.width = width


def generate(template, plan, output):
    prs = Presentation(template)
    # Reuse slide layouts so every new slide retains the branded masters and placeholders.
    # Capture placeholder indices: new slides have different shape IDs from the originals.
    original_placeholder_idx = {number: {sh.shape_id: sh.placeholder_format.idx
                                  for sh in prs.slides[number-1].shapes if sh.is_placeholder}
                                for number in (3, 4, 5)}
    # Remove the three sample content slides; retain cover and Slide 2.
    slide_ids = prs.slides._sldIdLst
    for index in (4, 3, 2):
        node = slide_ids[index]
        prs.part.drop_rel(node.rId)
        slide_ids.remove(node)
    # Locate layouts by names from inspected master template.
    names = {3: 'KV_Crops_V1', 4: 'KV Master only 1:1', 5: 'Product Film'}
    resolved = {number: next((l for l in prs.slide_layouts if l.name == name), None)
                for number, name in names.items()}
    if any(v is None for v in resolved.values()):
        raise ValueError('Required template slide layouts not found: ' + str(names))
    with tempfile.TemporaryDirectory() as td:
        temp = Path(td)
        for index, page in enumerate(plan, 1):
            slide = prs.slides.add_slide(resolved[page['layout']])
            def sid(original_id):
                idx = original_placeholder_idx[page['layout']][original_id]
                return next(sh.shape_id for sh in slide.shapes
                            if sh.is_placeholder and sh.placeholder_format.idx == idx)
            put_title(slide, page['instrument'])
            if page['layout'] == 3:
                for ratio, r in page['slots'].items():
                    put_image(slide, sid(IMAGE_SLOTS[ratio]), r['path'])
                put_table(slide, sid(8), page['assets'], 3)
            elif page['layout'] == 4:
                put_image(slide, sid(3), page['assets'][0]['path'])
                put_table(slide, sid(8), page['assets'], 4)
            else:
                video_path = page['assets'][0]['path']
                print(f'Processing video: {video_path.name}', flush=True)
                put_video(slide, sid(6), video_path, temp)
                put_table(slide, sid(7), page['assets'], 5)
            print(f'[{index}/{len(plan)}] {page["instrument"]}: {page["layout"]} ({len(page["assets"])} assets)', flush=True)
        prs.save(output)
    print('Saved:', output)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--plan-only', action='store_true', help='Validate Excel join and plan without media files')
    args = parser.parse_args()
    root = project_root()
    inp = root / 'input'
    matched_path = inp / 'assets_table1_cleaned.xlsx'
    raw = inp / 'Export_report_20260917_113237_Raw(1).xlsx'
    template = inp / 'master_template.pptx'
    for p in (excel, raw, template):
        if not p.is_file():
            raise FileNotFoundError(f'Missing input: {p}')
    records = read_assets(excel, raw)
    if args.plan_only:
        for r in records:
            r['path'] = Path(r['name'])
    else:
        find_local_files(records, inp / 'Assets')
    missing = [r for r in records if not r.get('path')]
    if missing:
        print(f'WARNING: {len(missing)} assets could not be located unambiguously')
        for r in missing:
            print(f'  Excel row {r["index"]}: {r["name"]} ({r.get("problem")})')
    plan = group_and_allocate(records)
    manifest = root / 'automation_manifest.csv'                             #create csv file to merge and join the 2 excel files (raw and cleaned) to get all metadata
    with manifest.open('w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['output_slide', 'template_slide', 'instrument', 'asset_id', 'file_name', 'ratio', 'local_path'])
        for slide_no, p in enumerate(plan, 3):
            for r in p['assets']:
                writer.writerow([slide_no, p['layout'], p['instrument'], r['id'], r['name'], r['ratio'], r['path']])
    print(f'Excel assets: {len(records)} | Found: {len(records)-len(missing)} | Planned slides: {len(plan)}')
    print('Manifest:', manifest)
    if not args.plan_only:
        generate(template, plan, root / 'final_output.pptx')


if __name__ == '__main__':
    main()
