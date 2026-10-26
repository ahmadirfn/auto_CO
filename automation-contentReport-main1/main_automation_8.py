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
from copy import deepcopy
import hashlib
import json
import math
import re
import shutil
import subprocess
import tempfile
import time
from collections import defaultdict
from pathlib import Path

import pandas as pd
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import CONTENT_TYPE, RELATIONSHIP_TYPE
from pptx.opc.packuri import PackURI
from pptx.parts.slide import SlideLayoutPart
from pptx.util import Cm, Pt
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn

IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tif', '.tiff', '.webp'}
VIDEO_EXTS = {'.mp4', '.mov', '.m4v', '.avi', '.wmv', '.webm', '.mkv'}
IMAGE_SLOTS = {'9:16': 3, '16:9': 11, '4:3': 13, '4:5': 12}
TABLE_HEADERS = ['Asset ID', 'Color Space', 'Aspect Ratio', 'Format', 'Symphony']
CORPORATE_FONT = 'MB Corpo S Text Office Light'
TABLE_HEADER_HEIGHT = Cm(1)
TABLE_DATA_MIN_HEIGHT = Cm(0.8)
COMPRESSION_ENABLED = True
MAX_COMPRESS_BYTES = 1_500_000
TARGET_MIN_BYTES = 1_200_000
JPEG_QUALITY_START = 95
JPEG_QUALITY_MIN = 80
JPEG_QUALITY_STEP = 5
VIDEO_CRF_START = 23
VIDEO_CRF_MAX = 25
VIDEO_PRESET = 'slow'
CACHE_VERSION = 1
PERFORMANCE_SECONDS = {}


def _record_timing(label, started):
    PERFORMANCE_SECONDS[label] = PERFORMANCE_SECONDS.get(label, 0.0) + (time.perf_counter() - started)


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
    left = shape.left + (shape.width-width)//2
    top = shape.top + (shape.height-height)//2
    if left < shape.left or top < shape.top or left + width > shape.left + shape.width or top + height > shape.top + shape.height:
        raise ValueError(f'Contained image exceeds placeholder bounds: {image_path}')
    return left, top, width, height


def _boxes_overlap(a, b):
    return a.left < b.left + b.width and b.left < a.left + a.width and \
        a.top < b.top + b.height and b.top < a.top + a.height


def sha256_file(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest()


def ffprobe_path():
    return shutil.which('ffprobe')


def probe_video(path):
    executable = ffprobe_path()
    if not executable:
        raise RuntimeError('FFprobe is not installed or is not available on PATH')
    result = subprocess.run([executable, '-v', 'error', '-show_streams', '-show_format',
                             '-of', 'json', str(path)], capture_output=True, text=True,
                            timeout=60, check=True)
    return json.loads(result.stdout)


def validate_video(source, output):
    original, encoded = probe_video(source), probe_video(output)
    original_v = next(s for s in original['streams'] if s.get('codec_type') == 'video')
    output_v = next(s for s in encoded['streams'] if s.get('codec_type') == 'video')
    for key in ('width', 'height'):
        if original_v.get(key) != output_v.get(key):
            raise ValueError(f'Compressed video changed {key}')
    for key in ('avg_frame_rate', 'r_frame_rate'):
        a, b = original_v.get(key), output_v.get(key)
        if a and b and a not in ('0/0', 'N/A') and b not in ('0/0', 'N/A'):
            def number(rate):
                n, d = rate.split('/')
                return float(n) / float(d)
            if abs(number(a) - number(b)) > 0.01:
                raise ValueError('Compressed video changed frame rate')
    duration_a = float(original.get('format', {}).get('duration', 0) or 0)
    duration_b = float(encoded.get('format', {}).get('duration', 0) or 0)
    if duration_a and abs(duration_a-duration_b) > max(0.5, duration_a * 0.02):
        raise ValueError('Compressed video duration differs from source')
    has_audio = any(s.get('codec_type') == 'audio' for s in original.get('streams', []))
    if has_audio and not any(s.get('codec_type') == 'audio' for s in encoded.get('streams', [])):
        raise ValueError('Compressed video lost audio')
    subprocess.run([ffmpeg_path(), '-v', 'error', '-i', str(output), '-f', 'null', '-'],
                   capture_output=True, text=True, timeout=180, check=True)


def compress_image(source, destination):
    from PIL import ImageOps
    ext = source.suffix.lower()
    with Image.open(source) as original:
        image = ImageOps.exif_transpose(original)
        image.load()
        if ext in {'.jpg', '.jpeg'}:
            if image.mode not in {'RGB', 'L'}:
                image = image.convert('RGB')
            kwargs = {'format': 'JPEG', 'optimize': True}
            if original.info.get('icc_profile'):
                kwargs['icc_profile'] = original.info['icc_profile']
            try:
                exif = original.getexif()
                if exif:
                    exif[274] = 1
                    kwargs['exif'] = exif.tobytes()
            except (AttributeError, TypeError):
                pass
            best_quality = JPEG_QUALITY_START
            best_size = None
            for quality in range(JPEG_QUALITY_START, JPEG_QUALITY_MIN - 1, -JPEG_QUALITY_STEP):
                image.save(destination, quality=quality, **kwargs)
                size = destination.stat().st_size
                if best_size is None or size < best_size:
                    best_size, best_quality = size, quality
                if size <= MAX_COMPRESS_BYTES:
                    best_quality = quality
                    break
            image.save(destination, quality=best_quality, **kwargs)
            return 'Pillow JPEG optimize', f'quality {best_quality}'
        kwargs = {'format': 'PNG', 'optimize': True}
        if original.info.get('icc_profile'):
            kwargs['icc_profile'] = original.info['icc_profile']
        try:
            exif = original.getexif()
            if exif:
                exif[274] = 1
                kwargs['exif'] = exif.tobytes()
        except (AttributeError, TypeError):
            pass
        image.save(destination, **kwargs)
        with Image.open(destination) as check:
            check.verify()
        return 'Pillow PNG lossless optimize', 'lossless'


def compress_video(source, destination):
    ffmpeg = ffmpeg_path()
    original = probe_video(source)
    best = None
    for crf in range(VIDEO_CRF_START, VIDEO_CRF_MAX + 1, 2):
        temp_output = destination.with_name(destination.stem + f'.crf{crf}.tmp.mp4')
        temp_output.unlink(missing_ok=True)
        command = [ffmpeg, '-nostdin', '-hide_banner', '-loglevel', 'error', '-y', '-i', str(source),
                   '-map', '0:v:0', '-map', '0:a?', '-c:v', 'libx264', '-preset', VIDEO_PRESET,
                   '-crf', str(crf), '-pix_fmt', 'yuv420p', '-c:a', 'aac', '-b:a', '128k',
                   '-movflags', '+faststart', str(temp_output)]
        try:
            print(f'Compressing video {source.name} (CRF {crf})...', flush=True)
            subprocess.run(command, capture_output=True, text=True, timeout=900, check=True)
            validate_video(source, temp_output)
            size = temp_output.stat().st_size
            if best:
                best[1].unlink(missing_ok=True)
            best = (size, temp_output, crf)
            if size <= MAX_COMPRESS_BYTES:
                break
        except Exception:
            temp_output.unlink(missing_ok=True)
            if not best:
                raise
    if not best:
        raise RuntimeError('FFmpeg did not produce a valid compressed output')
    best[1].replace(destination)
    return 'FFmpeg H.264/AAC', f'CRF {best[2]}'


def load_cache(path):
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return {}


def compress_selected_media(plan, assets_root, compressed_root):
    cache_file = compressed_root / 'compression_cache.json'
    cache = load_cache(cache_file)
    settings = {'version': CACHE_VERSION, 'max_bytes': MAX_COMPRESS_BYTES, 'jpeg': [JPEG_QUALITY_START, JPEG_QUALITY_MIN, JPEG_QUALITY_STEP],
                'crf': [VIDEO_CRF_START, VIDEO_CRF_MAX], 'preset': VIDEO_PRESET}
    settings_key = hashlib.sha256(json.dumps(settings, sort_keys=True).encode()).hexdigest()
    report, selected_by_path = [], {}
    supported = {'.jpg', '.jpeg', '.png', '.mp4', '.mov'}
    for page in plan:
        for asset in page['assets']:
            source = Path(asset['path'])
            asset['selected_media_path'] = source
            source_key = str(source.resolve())
            if source_key in selected_by_path:
                asset['selected_media_path'] = selected_by_path[source_key]
                continue
            size = source.stat().st_size
            row = {'original_path': str(source), 'selected_path': str(source), 'file_name': source.name,
                   'media_type': 'Image' if source.suffix.lower() in IMAGE_EXTS else 'Video',
                   'format': source.suffix.lstrip('.').upper(), 'original_size_mb': size / 1e6,
                   'selected_size_mb': size / 1e6, 'reduction_mb': 0.0, 'reduction_percent': 0.0,
                   'compression_status': 'Skipped', 'compression_method': '', 'quality_setting': '',
                   'cache_hit': False, 'notes': ''}
            if not COMPRESSION_ENABLED or size <= MAX_COMPRESS_BYTES:
                row['notes'] = 'At or below 1.5 MB' if COMPRESSION_ENABLED else 'Compression disabled'
            elif source.suffix.lower() not in supported:
                row['notes'] = 'Unsupported format'
            else:
                try:
                    rel = source.resolve().relative_to(assets_root.resolve())
                    cache_key = str(rel).replace('\\', '/')
                    digest = sha256_file(source)
                    entry = cache.get(cache_key, {})
                    cached_rel = entry.get('compressed_path')
                    cached = compressed_root / cached_rel if cached_rel else None
                    cached_size = cached.stat().st_size if cached and cached.is_file() else 0
                    recorded_size = entry.get('compressed_size')
                    entry_matches = (entry.get('sha256') == digest and entry.get('settings') == settings_key)
                    rejected_cache_hit = entry_matches and entry.get('status') == 'Rejected'
                    cache_valid = (entry_matches
                                   and cached and cached.is_file() and cached_size > 0
                                   and (recorded_size is None or cached_size == recorded_size))
                    if rejected_cache_hit:
                        selected = source
                        row.update(compression_status='Rejected', cache_hit=True,
                                   notes='Cached result: compression did not reduce size')
                    elif cache_valid:
                        selected = cached
                        row.update(compression_status='Reused', compression_method=entry.get('method', ''),
                                   quality_setting=entry.get('quality', ''), cache_hit=True)
                    else:
                        if source.suffix.lower() == '.mov':
                            collision_tag = hashlib.sha256(str(rel).encode()).hexdigest()[:8]
                            output_rel = rel.with_name(f'{rel.stem}__mov_{collision_tag}.mp4')
                        else:
                            output_rel = rel
                        selected = compressed_root / output_rel
                        selected.parent.mkdir(parents=True, exist_ok=True)
                        with tempfile.TemporaryDirectory(dir=selected.parent, prefix='.compress-') as temp_dir:
                            temp_output = Path(temp_dir) / selected.name
                            if source.suffix.lower() in {'.jpg', '.jpeg', '.png'}:
                                method, quality = compress_image(source, temp_output)
                            else:
                                method, quality = compress_video(source, temp_output)
                            out_size = temp_output.stat().st_size
                            if out_size >= size:
                                selected = source
                                row.update(compression_status='Rejected', notes='Compression did not reduce size')
                                cache[cache_key] = {'sha256': digest, 'source_size': size,
                                                    'source_mtime_ns': source.stat().st_mtime_ns,
                                                    'compressed_path': '', 'compressed_size': size,
                                                    'settings': settings_key, 'method': method,
                                                    'quality': quality, 'status': 'Rejected'}
                            else:
                                if source.suffix.lower() in {'.mp4', '.mov'}:
                                    validate_video(source, temp_output)
                                temp_output.replace(compressed_root / output_rel)
                                selected = compressed_root / output_rel
                                row.update(compression_status='Compressed', compression_method=method,
                                           quality_setting=quality)
                                cache[cache_key] = {'sha256': digest, 'source_size': size,
                                                    'source_mtime_ns': source.stat().st_mtime_ns,
                                                    'compressed_path': str(output_rel).replace('\\', '/'),
                                                    'compressed_size': out_size, 'settings': settings_key,
                                                    'method': method, 'quality': quality, 'status': 'Compressed'}
                    row['selected_path'] = str(selected)
                    row['selected_size_mb'] = selected.stat().st_size / 1e6
                    row['reduction_mb'] = row['original_size_mb'] - row['selected_size_mb']
                    row['reduction_percent'] = (100 * row['reduction_mb'] / row['original_size_mb']) if size else 0
                    asset['selected_media_path'] = selected
                except Exception as exc:
                    row.update(compression_status='Failed', notes=str(exc))
                    print(f'WARNING: {source.name}: compression failed ({exc}); using original', flush=True)
            selected_by_path[source_key] = asset['selected_media_path']
            report.append(row)
    compressed_root.mkdir(parents=True, exist_ok=True)
    cache_file.write_text(json.dumps(cache, indent=2), encoding='utf-8')
    report_path = assets_root.parent.parent / 'compression_report.csv'
    fields = ['original_path', 'selected_path', 'file_name', 'media_type', 'format', 'original_size_mb',
              'selected_size_mb', 'reduction_mb', 'reduction_percent', 'compression_status',
              'compression_method', 'quality_setting', 'cache_hit', 'notes']
    with report_path.open('w', newline='', encoding='utf-8-sig') as stream:
        writer = csv.DictWriter(stream, fieldnames=fields)
        writer.writeheader()
        for row in report:
            writer.writerow({k: f'{row[k]:.4f}' if k in {'original_size_mb', 'selected_size_mb', 'reduction_mb', 'reduction_percent'} else row[k] for k in fields})
    counts = {status: sum(r['compression_status'] == status for r in report)
              for status in ('Compressed', 'Skipped', 'Reused', 'Failed', 'Rejected')}
    original_total = sum(r['original_size_mb'] for r in report)
    selected_total = sum(r['selected_size_mb'] for r in report)
    print(f"Compression: {counts['Compressed']} compressed, {counts['Skipped']} skipped, {counts['Reused']} reused, {counts['Failed']} failed, {counts['Rejected']} rejected; size reduction {original_total-selected_total:.2f} MB")
    print(f'Compression report: {report_path}')
    return report


def put_image(slide, shape_id, path):
    sh = find_shape(slide, shape_id)
    x, y, w, h = contain_rect(sh, path)
    pic = slide.shapes.add_picture(str(path), x, y, width=w, height=h)
    pic.name = 'Generated media image'
    replace_shape(sh, pic)
    return pic


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


def make_poster(video, destination, cache_file=None):
    started = time.perf_counter()
    if cache_file is not None:
        cache = load_cache(cache_file)
        thumbnail_cache = cache.setdefault('__thumbnails__', {})
        source_key = str(Path(video).resolve())
        source_digest = sha256_file(video)
        thumbnail_settings = {'version': 1, 'timestamps': ['2', '0'], 'frames': 1, 'format': 'jpg'}
        settings_key = hashlib.sha256(json.dumps(thumbnail_settings, sort_keys=True).encode()).hexdigest()
        key = hashlib.sha256(source_key.encode('utf-8')).hexdigest()
        entry = thumbnail_cache.get(key, {})
        cached_path = cache_file.parent / entry.get('thumbnail_path', '')
        if (entry.get('source_sha256') == source_digest and entry.get('settings') == settings_key
                and cached_path.is_file() and cached_path.stat().st_size > 0):
            _record_timing('Video thumbnail generation/reuse', started)
            return cached_path
        destination = cache_file.parent / 'thumbnails' / f'{Path(video).stem}_{source_digest[:16]}.jpg'
        destination.parent.mkdir(parents=True, exist_ok=True)
    # Try the 2-second frame first, then the first frame for short videos.
    # A timeout prevents one problematic file from blocking the entire run.
    try:
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
                if cache_file is not None:
                    thumbnail_cache[key] = {
                        'source_sha256': source_digest,
                        'settings': settings_key,
                        'thumbnail_path': str(destination.relative_to(cache_file.parent)).replace('\\', '/'),
                        'thumbnail_size': destination.stat().st_size,
                    }
                    cache_file.write_text(json.dumps(cache, indent=2), encoding='utf-8')
                return destination
            print(f'WARNING: Thumbnail failed at {timestamp}s: {video.name} '
                  f'{result.stderr[-250:]}', flush=True)
        raise RuntimeError(f'Unable to extract thumbnail from: {video}')
    finally:
        _record_timing('Video thumbnail generation/reuse', started)


def put_video(slide, shape_id, path, temp, thumbnail_cache_file=None):
    sh = find_shape(slide, shape_id)
    poster = temp / (asset_id(path.name) + '_poster.jpg')
    poster = make_poster(path, poster, thumbnail_cache_file)
    # Fit the video inside its placeholder, without stretching or cropping.
    with Image.open(poster) as im:
        width_px, height_px = im.size
    factor = min(sh.width / width_px, sh.height / height_px)
    width, height = int(width_px * factor), int(height_px * factor)
    left = sh.left + (sh.width - width) // 2
    top = sh.top + (sh.height - height) // 2
    movie = slide.shapes.add_movie(str(path), left, top, width, height,
                                   poster_frame_image=str(poster), mime_type='video/mp4')
    movie.name = 'Generated media video'
    replace_shape(sh, movie)


def put_title(slide, instrument):
    sh = find_shape(slide, 2)
    if sh.has_text_frame:
        sh.text = instrument
        for paragraph in sh.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(17 if len(instrument) > 48 else 24)


def suppress_metadata_headers(prs, layout_overrides=None):
    """Remove only inherited metadata labels from the three generated content layouts."""
    names = {
        'KV_Crops_V1': {'Asset ID', 'Color Space', 'Aspect Ratio', 'Format', 'Symphony Link'},
        'KV Master only 1:1': {'Asset ID', 'Color Space', 'Aspect Ratio', 'Format', 'Symphony Link'},
        'Product Film': {'Asset ID', 'Variation', 'Quality', 'Symphony Link'},
    }
    header_top = {}
    layout_overrides = layout_overrides or {}
    for layout_name, labels in names.items():
        layout = layout_overrides.get(layout_name) or next(
            (item for item in prs.slide_layouts if item.name == layout_name), None)
        if layout is None:
            continue
        matches = [shape for shape in layout.shapes
                   if shape.has_text_frame and shape.text.strip() in labels]
        if matches:
            header_top[layout_name] = min(shape.top for shape in matches)
        for shape in matches:
            shape._element.getparent().remove(shape._element)
    return header_top


def clean_picture_library_v2_layout(layout):
    """Hide inherited V2 labels and its unused off-slide table in this output only."""
    if layout is None:
        return
    labels = {'Asset ID', 'Color Space', 'Aspect Ratio', 'Format', 'Symphony Link'}
    for shape in list(layout.shapes):
        if shape.is_placeholder and shape.placeholder_format.idx == 17:
            shape._element.getparent().remove(shape._element)
        elif shape.has_text_frame and shape.text.strip() in labels:
            shape._element.getparent().remove(shape._element)


def _clear_cell_decoration(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    for child in list(tc_pr):
        if child.tag in {qn('a:solidFill'), qn('a:gradFill'), qn('a:blipFill'), qn('a:lnL'),
                         qn('a:lnR'), qn('a:lnT'), qn('a:lnB')}:
            tc_pr.remove(child)
    no_fill = OxmlElement('a:noFill')
    tc_pr.append(no_fill)
    for edge in ('L', 'R', 'T', 'B'):
        line = OxmlElement(f'a:ln{edge}')
        line.set('w', '0')
        line.append(OxmlElement('a:noFill'))
        tc_pr.append(line)


def _estimate_lines(value, width_emu, font_pt):
    # Conservative character-width estimate; descriptive text wraps while compact IDs stay single-line.
    width_pt = max(1.0, width_emu / 12700 - 6)  # subtract cell padding
    capacity = max(3, int(width_pt / max(font_pt * 0.52, 1)))
    lines = 0
    for paragraph in str(value).splitlines() or ['']:
        words = paragraph.split()
        if not words:
            lines += 1
            continue
        current = 0
        for word in words:
            length = len(word)
            if current == 0:
                used_lines = max(1, math.ceil(length / capacity))
                current = length % capacity or capacity
                lines += used_lines
            elif current + 1 + length <= capacity:
                current += 1 + length
            else:
                current = min(length, capacity)
                lines += 1
    return max(1, lines)


def _dynamic_column_widths(headers, rows, total_width):
    # Allocate widths from header and cell content, with a useful minimum for compact metadata columns.
    mins_cm = [1.45 if 'Asset ID' in headers[0] else 1.0 for _ in headers]
    for i, label in enumerate(headers):
        if label in {'Color Space', 'Aspect Ratio', 'Format', 'Quality'}:
            mins_cm[i] = 1.35
        elif label == 'Symphony':
            mins_cm[i] = 1.75
        elif label in {'Variation', 'Description'}:
            mins_cm[i] = 2.0
    minima = [Cm(v) for v in mins_cm]
    if sum(minima) > total_width:
        minima = [int(total_width / len(headers))] * len(headers)
    remaining = total_width - sum(minima)
    weights = []
    for ci, header in enumerate(headers):
        longest = max([len(header)] + [len(str(row[ci])) for row in rows])
        # Cap long descriptions so they wrap instead of stealing all width from other columns.
        weights.append(max(1.0, min(longest, 32) ** 0.72))
    weight_sum = sum(weights)
    widths = [minima[i] + int(remaining * weights[i] / weight_sum) for i in range(len(headers))]
    widths[-1] += total_width - sum(widths)
    return widths


def _description_left(slide, layout, fallback):
    description_idx = next((shape.placeholder_format.idx for shape in layout.shapes
                            if shape.is_placeholder and shape.has_text_frame
                            and shape.text.strip().strip('[]').casefold() == 'description'), None)
    if description_idx is not None:
        return next((shape.left for shape in slide.shapes if shape.is_placeholder
                     and shape.placeholder_format.idx == description_idx), fallback)
    return fallback


def put_table(slide, shape_id, assets, layout_number, header_top, layout_obj):
    placeholder = find_shape(slide, shape_id)
    video_layout = layout_number == 5
    headers = ['Asset ID', 'Variation', 'Quality', 'Symphony'] if video_layout else TABLE_HEADERS
    data = [headers]
    links = {}
    for ri, r in enumerate(assets, start=1):
        info = r['raw']
        link = str(info.get('Hyperlink', '')).strip()
        if video_layout:
            row = [r['id'], str(info.get('Perspective', '')).strip() or '[Variation]',
                   '[Quality]', 'Link' if link else '[Link]']
        else:
            row = [r['id'], str(info.get('ColorSpace', '')).strip() or '[Color Space]', r['ratio'],
                   Path(r['path']).suffix.lstrip('.').upper(),
                   'Link' if link else '[Link]']
        if link.startswith(('https://', 'http://')):
            links[(ri, len(headers)-1)] = link
        data.append(row)

    total_width = placeholder.width
    left = _description_left(slide, layout_obj, placeholder.left)
    top = header_top if header_top is not None else placeholder.top - Cm(0.56)
    widths = _dynamic_column_widths(headers, data[1:], total_width)
    font_pt = 8.0 if layout_number == 3 else 9.0
    available_height = placeholder.top + placeholder.height - top
    while True:
        row_heights = [TABLE_HEADER_HEIGHT]
        for row in data[1:]:
            line_count = max(_estimate_lines(value, widths[i], font_pt) for i, value in enumerate(row))
            row_heights.append(max(TABLE_DATA_MIN_HEIGHT, Cm(0.8 * line_count)))
        if sum(row_heights) <= available_height or font_pt <= 7.0:
            break
        font_pt -= 0.5
    total_height = sum(row_heights)
    if total_height > available_height:
        raise ValueError(f'Metadata table does not fit its template area on slide {slide.slide_id}; split the content across slides.')
    table_bounds = type('Bounds', (), {'left': left, 'top': top, 'width': total_width, 'height': total_height})()
    for media_shape in slide.shapes:
        if media_shape.name.startswith('Generated media ') and _boxes_overlap(media_shape, table_bounds):
            raise ValueError(f'Metadata table overlaps {media_shape.name} on slide {slide.slide_id}')
    frame = slide.shapes.add_table(len(data), len(headers), left, top, total_width, total_height)
    replace_shape(placeholder, frame)
    table = frame.table
    for column, width in zip(table.columns, widths):
        column.width = width
    for ri, height in enumerate(row_heights):
        table.rows[ri].height = height
    for ri, row in enumerate(data):
        for ci, value in enumerate(row):
            cell = table.cell(ri, ci)
            cell.text = str(value)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Pt(3)
            cell.margin_top = cell.margin_bottom = Pt(1)
            _clear_cell_decoration(cell)
            cell.text_frame.word_wrap = ri > 0 and (
                headers[ci] in {'Variation', 'Description'} or
                (headers[ci] == 'Symphony' and len(str(value)) > 18)
            )
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER
                paragraph.space_before = Pt(0)
                paragraph.space_after = Pt(0)
                for run in paragraph.runs:
                    run.font.name = CORPORATE_FONT
                    run.font.size = Pt(font_pt)
                    run.font.bold = ri == 0
                    if ri == 0:
                        run.font.color.rgb = RGBColor(0, 0, 0)
                    if (ri, ci) in links:
                        run.hyperlink.address = links[(ri, ci)]


def put_picture_library_v2_table(slide, shape_id, asset, left=None):
    """Replace one V2 metadata area with a five-row label/value table."""
    placeholder = find_shape(slide, shape_id)
    info = asset['raw']
    link = str(info.get('Hyperlink', '')).strip()
    data = [
        ('Asset ID', asset['id']),
        ('Color Space', str(info.get('ColorSpace', '')).strip() or '[Color Space]'),
        ('Aspect Ratio', asset['ratio']),
        ('Format', Path(asset['path']).suffix.lstrip('.').upper()),
        ('Symphony', 'Link' if link else '[Link]'),
    ]
    table_left = placeholder.left if left is None else left
    bounds = type('Bounds', (), {'left': table_left, 'top': placeholder.top,
                                 'width': placeholder.width, 'height': placeholder.height})()
    for media_shape in slide.shapes:
        if media_shape.name.startswith('Generated media ') and _boxes_overlap(media_shape, bounds):
            raise ValueError(f'Picture Library V2 metadata overlaps {media_shape.name}')
    frame = slide.shapes.add_table(len(data), 2, table_left, placeholder.top,
                                   placeholder.width, placeholder.height)
    replace_shape(placeholder, frame)
    table = frame.table
    label_width = placeholder.width // 2
    table.columns[0].width = label_width
    table.columns[1].width = placeholder.width - label_width
    row_height = placeholder.height // len(data)
    for ri, (label, value) in enumerate(data):
        table.rows[ri].height = row_height if ri < len(data) - 1 else placeholder.height - row_height * (len(data) - 1)
        for ci, text in enumerate((label, value)):
            cell = table.cell(ri, ci)
            cell.text = str(text)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Pt(2)
            cell.margin_top = cell.margin_bottom = Pt(0)
            _clear_cell_decoration(cell)
            cell.text_frame.word_wrap = ci == 1 and ri in (1, 4)
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.LEFT
                paragraph.space_before = Pt(0)
                paragraph.space_after = Pt(0)
                for run in paragraph.runs:
                    run.font.name = CORPORATE_FONT
                    run.font.size = Pt(8)
                    run.font.bold = ci == 0
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    if ci == 1 and ri == 4 and link.startswith(('https://', 'http://')):
                        run.hyperlink.address = link


def populate_picture_library_v2(slide, page, placeholder_indices):
    def shape_id_for_idx(idx):
        if idx not in placeholder_indices:
            raise ValueError(f'KV_Crops_V2 template is missing placeholder index {idx}')
        return next(shape.shape_id for shape in slide.shapes
                    if shape.is_placeholder and shape.placeholder_format.idx == idx)

    put_title(slide, page['instrument'])
    # Template geometry: idx 22 is landscape (4:3); idx 21 is portrait (4:5).
    picture_idx = {'9:16': 13, '16:9': 20, '4:3': 22, '4:5': 21}
    table_idx = {'16:9': 23, '9:16': 24, '4:3': 26, '4:5': 25}
    active_tables = {table_idx[ratio] for ratio in page['slots']}
    for shape in list(slide.shapes):
        if shape.is_placeholder and shape.placeholder_format.idx in ({17} | (set(table_idx.values()) - active_tables)):
            shape._element.getparent().remove(shape._element)
    media_by_ratio = {}
    for ratio, asset in page['slots'].items():
        picture_shape_id = shape_id_for_idx(picture_idx[ratio])
        picture_left = find_shape(slide, picture_shape_id).left
        media_by_ratio[ratio] = put_image(
            slide, picture_shape_id,
            asset.get('selected_media_path', asset['path']))
        put_picture_library_v2_table(slide, shape_id_for_idx(table_idx[ratio]), asset,
                                     left=picture_left)
    # Align the visible lower edges of the 4:3 and 4:5 pictures with the 9:16 picture.
    reference = media_by_ratio.get('9:16')
    if reference is not None:
        reference_bottom = reference.top + reference.height
        for ratio in ('4:3', '4:5'):
            media = media_by_ratio.get(ratio)
            if media is not None and media.top + media.height < reference_bottom:
                media.top += reference_bottom - (media.top + media.height)
    return slide


def center_picture_library_v2_labels(layout):
    labels = {'Asset ID', 'Color Space', 'Aspect Ratio', 'Format', 'Symphony Link'}
    for shape in layout.shapes:
        if shape.has_text_frame and shape.text.strip() in labels:
            for paragraph in shape.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER


def _clone_layout(prs, source_layout):
    """Create a private layout copy for populated slides, leaving blank templates intact."""
    package = prs.part.package
    used_numbers = []
    for part in package.iter_parts():
        match = re.fullmatch(r'/ppt/slideLayouts/slideLayout(\d+)\.xml', str(part.partname))
        if match:
            used_numbers.append(int(match.group(1)))
    partname = PackURI(f'/ppt/slideLayouts/slideLayout{max(used_numbers, default=0) + 1}.xml')
    clone = SlideLayoutPart(partname, CONTENT_TYPE.PML_SLIDE_LAYOUT, package,
                            deepcopy(source_layout.part._element))
    for rel in source_layout.part.rels.values():
        if rel.is_external:
            clone.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            clone.rels.get_or_add(rel.reltype, rel.target_part)
    master_part = source_layout.part.slide_master.part
    layout_rid = master_part.relate_to(clone, RELATIONSHIP_TYPE.SLIDE_LAYOUT)
    layout_ids = master_part._element.sldLayoutIdLst
    new_id = max((int(item.get('id')) for item in layout_ids), default=0) + 1
    entry = OxmlElement('p:sldLayoutId')
    entry.set('id', str(new_id))
    entry.set(qn('r:id'), layout_rid)
    layout_ids.append(entry)
    return clone.slide_layout


def _duplicate_template_slide(prs, source_slide, layout):
    """Duplicate a template slide's complete slide XML and relationships."""
    duplicate = prs.slides.add_slide(layout)
    rel_ids = {}
    for rel in source_slide.part.rels.values():
        if rel.reltype == RELATIONSHIP_TYPE.SLIDE_LAYOUT:
            new_rid = duplicate.part.rels.get_or_add(RELATIONSHIP_TYPE.SLIDE_LAYOUT, layout.part)
        elif rel.is_external:
            new_rid = duplicate.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            new_rid = duplicate.part.rels.get_or_add(rel.reltype, rel.target_part)
        rel_ids[rel.rId] = new_rid
    slide_xml = deepcopy(source_slide.part._element)
    relationship_attributes = {qn('r:id'), qn('r:embed'), qn('r:link')}
    for element in slide_xml.iter():
        for attribute in relationship_attributes:
            old_rid = element.get(attribute)
            if old_rid in rel_ids:
                element.set(attribute, rel_ids[old_rid])
    duplicate.part._element = slide_xml
    duplicate._element = slide_xml
    duplicate.__dict__.pop('shapes', None)
    return duplicate


def generate(template, plan, output, thumbnail_cache_file=None):
    slide_started = time.perf_counter()
    prs = Presentation(template)
    cover_slide = prs.slides[0]
    master_layouts = [layout for master in prs.slide_masters for layout in master.slide_layouts]
    # These are the actual template slides whose slide-level shapes are duplicated.
    sources = {2: prs.slides[1], 3: prs.slides[2], 4: prs.slides[3], 5: prs.slides[4]}
    original_slide_ids = {number: {sh.shape_id: sh.placeholder_format.idx
                                   for sh in sources[number].shapes if sh.is_placeholder}
                          for number in (3, 4, 5)}
    v2_placeholder_indices = {sh.placeholder_format.idx for sh in sources[2].shapes if sh.is_placeholder}

    names = {2: 'KV_Crops_V2', 3: 'KV_Crops_V1',
             4: 'KV Master only 1:1', 5: 'Product Film'}
    source_layouts = {number: next((layout for layout in prs.slide_layouts if layout.name == name), None)
                      for number, name in names.items()}
    if any(layout is None for layout in source_layouts.values()):
        raise ValueError('Required template slide layouts not found: ' + str(names))

    # A layout-3 plan page produces both existing Picture Library arrangements.
    jobs = []
    for page in plan:
        if page['layout'] == 3:
            jobs.extend(((2, page, 3, 'v2'), (3, page, 3, 'v1')))
        else:
            jobs.append((page['layout'], page, page['layout'], 'standard'))

    populated_layouts = {number: _clone_layout(prs, layout)
                         for number, layout in source_layouts.items()}
    metadata_header_top = suppress_metadata_headers(
        prs, layout_overrides={
            'KV_Crops_V1': populated_layouts[3],
            'KV Master only 1:1': populated_layouts[4],
            'Product Film': populated_layouts[5],
        })
    clean_picture_library_v2_layout(populated_layouts[2])

    def remove_slide(slide):
        for node in list(prs.slides._sldIdLst):
            rel = prs.part.rels.get(node.rId)
            if rel is not None and rel.target_part is slide.part:
                prs.part.drop_rel(node.rId)
                prs.slides._sldIdLst.remove(node)
                return

    def populate(slide, page, layout_number, variant, temp):
        def sid(original_id):
            idx = original_slide_ids[layout_number][original_id]
            return next(sh.shape_id for sh in slide.shapes
                        if sh.is_placeholder and sh.placeholder_format.idx == idx)
        if variant == 'v2':
            populate_picture_library_v2(slide, page, v2_placeholder_indices)
        else:
            put_title(slide, page['instrument'])
            if layout_number == 3:
                for ratio, record in page['slots'].items():
                    put_image(slide, sid(IMAGE_SLOTS[ratio]),
                              record.get('selected_media_path', record['path']))
                put_table(slide, sid(8), page['assets'], 3,
                          metadata_header_top.get('KV_Crops_V1'), slide.slide_layout)
            elif layout_number == 4:
                record = page['assets'][0]
                put_image(slide, sid(3), record.get('selected_media_path', record['path']))
                put_table(slide, sid(8), page['assets'], 4,
                          metadata_header_top.get('KV Master only 1:1'), slide.slide_layout)
            else:
                video_path = page['assets'][0].get('selected_media_path', page['assets'][0]['path'])
                print(f'Processing video: {video_path.name}', flush=True)
                put_video(slide, sid(6), video_path, temp, thumbnail_cache_file)
                put_table(slide, sid(7), page['assets'], 5,
                          metadata_header_top.get('Product Film'), slide.slide_layout)
        print(f'{page["instrument"]}: {layout_number} ({len(page["assets"])} assets)', flush=True)

    with tempfile.TemporaryDirectory() as td:
        temp = Path(td)
        blank_layout_slides = [prs.slides.add_slide(layout) for layout in master_layouts]
        # The populated section begins with a copy of the original cover.
        cover_copy = _duplicate_template_slide(prs, cover_slide, cover_slide.slide_layout)
        populated_slides = []
        for source_number, page, layout_number, variant in jobs:
            slide = _duplicate_template_slide(prs, sources[source_number], populated_layouts[source_number])
            populate(slide, page, layout_number, variant, temp)
            populated_slides.append(slide)

        # Keep the untouched cover first; remove only the original sample content slides.
        for source in sources.values():
            remove_slide(source)
        expected = 2 + len(master_layouts) + len(jobs)
        if len(prs.slides) != expected:
            raise RuntimeError(f'Expected {expected} slides (two covers, all master layouts, and populated assets), got {len(prs.slides)}')
        ordered_slides = list(prs.slides)
        layout_start = 1
        populated_start = 1 + len(master_layouts) + 1
        if ordered_slides[0].part is not cover_slide.part:
            raise RuntimeError('Original cover slide is not first')
        if any(slide.slide_layout.part is not layout.part
               for slide, layout in zip(ordered_slides[layout_start:populated_start - 1], master_layouts)):
            raise RuntimeError('Blank template slide layout order does not match the Slide Masters')
        if any(slide.part is not expected_slide.part
               for slide, expected_slide in zip(ordered_slides[layout_start:populated_start - 1], blank_layout_slides)):
            raise RuntimeError('Blank template slides are not in the expected order')
        if ordered_slides[populated_start - 1].part is not cover_copy.part:
            raise RuntimeError('Cover copy is not the first slide in the populated section')
        if any(slide.part is not expected_slide.part
               for slide, expected_slide in zip(ordered_slides[populated_start:], populated_slides)):
            raise RuntimeError('Populated slides are not in the expected order')
        save_started = time.perf_counter()
        prs.save(output)
        _record_timing('Final PowerPoint save', save_started)
    _record_timing('Slide creation and population', slide_started)
    print(f'Saved: {output} (1 original cover, {len(master_layouts)} master-layout slides, 1 cover copy, {len(jobs)} populated slides; {2 + len(master_layouts) + len(jobs)} total)', flush=True)


def main():
    total_started = time.perf_counter()
    PERFORMANCE_SECONDS.clear()
    parser = argparse.ArgumentParser()
    parser.add_argument('--plan-only', action='store_true', help='Validate Excel join and plan without media files')
    args = parser.parse_args()
    root = project_root()
    inp = root / 'input'
    matched_path = inp / 'assets_table1_cleaned.xlsx'
    raw = inp / 'Export_report_20260917_113237_Raw(1).xlsx'
    template = inp / 'master_template.pptx'
    for p in (matched_path, raw, template):
        if not p.is_file():
            raise FileNotFoundError(f'Missing input: {p}')
    excel_started = time.perf_counter()
    records = read_assets(matched_path, raw)
    _record_timing('Excel reading and metadata matching', excel_started)
    if args.plan_only:
        for r in records:
            r['path'] = Path(r['name'])
    else:
        scan_started = time.perf_counter()
        find_local_files(records, inp / 'Assets')
        _record_timing('Asset directory scanning and matching', scan_started)
    missing = [r for r in records if not r.get('path')]
    if missing:
        print(f'WARNING: {len(missing)} assets could not be located unambiguously')
        for r in missing:
            print(f'  Excel row {r["index"]}: {r["name"]} ({r.get("problem")})')
    plan = group_and_allocate(records)
    if not args.plan_only:
        compression_started = time.perf_counter()
        compress_selected_media(plan, inp / 'Assets', inp / 'compressed_assets')
        _record_timing('Image/video compression and cache lookup', compression_started)
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
        generate(template, plan, root / 'final_output.pptx', inp / 'compressed_assets' / 'compression_cache.json')
    PERFORMANCE_SECONDS['Total execution'] = time.perf_counter() - total_started
    labels = (
        'Asset directory scanning and matching',
        'Excel reading and metadata matching',
        'Image/video compression and cache lookup',
        'Video thumbnail generation/reuse',
        'Slide creation and population',
        'Final PowerPoint save',
        'Total execution',
    )
    print('Performance summary:')
    for label in labels:
        print(f'  {label}: {PERFORMANCE_SECONDS.get(label, 0.0):.2f}s')


if __name__ == '__main__':
    main()
