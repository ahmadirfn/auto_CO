"""Enrich a Symphony export without replacing populated source metadata.

Usage: python map_content_overview.py [--input "export.xlsx"] [--output "mapped.xlsx"]
With no arguments, uses the raw export in the parent content_excel_mapping folder.
Requires: pip install openpyxl
"""
import argparse
import re
from pathlib import Path
from collections import Counter, defaultdict
from copy import copy

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

MAPPING_DIR = Path(__file__).resolve().parent.parent
DEFAULT_INPUT = MAPPING_DIR / 'Export_report_20260917_113237_Raw(1).xlsx'
DEFAULT_OUTPUT = MAPPING_DIR.parent / 'Content_Overview_Enriched.xlsx'

NEW = [
    'PPT Instrument', 'KV #', 'KV Description', 'KV Family', 'Film Duration',
    'Film Ratio', 'Edit Family', 'Variation', 'Asset Type', 'Quality',
    'Asset Role', 'Format', 'Group Key', 'Sort Order', 'Mapping Status',
    'Mapping Notes', 'Sequence #', 'Car Model', 'Motorization', 'Master',
    'Link to Stems',
]
RATIO = re.compile(r'(?<!\d)(16[-:]9|9[-:]16|1[-:]1|4[-:]5|5[-:]4|4[-:]3|3[-:]4|3[-:]2|2[-:]3|21[-:]9)(?!\d)', re.I)
DURATION = re.compile(r'(?<![A-Za-z0-9.])(?P<n>\d+(?:\.\d+)?)\s*(?:s(?:ec(?:onds?)?)?)(?![A-Za-z0-9])', re.I)
KV = re.compile(r'ProdStillUnreal_([0-9]+[a-z]?)[-_]', re.I)
KIMI = re.compile(r'(Keyvisual|Sideshot)_([0-9]+[a-z]?)(?:[-_]|$)', re.I)
SOCIAL_STILL = re.compile(r'SocialStill(?:[-_]|$)', re.I)
SOCIAL_FILM = re.compile(r'SocialFilm(?:[-_]|$)', re.I)
PROD_FILM = re.compile(r'ProdFilm(?:_|-|$)', re.I)
TVC = re.compile(r'(?:^|[-_])TVC(?:[-_]|$)', re.I)
SUPPORTED = re.compile(r'(?i)(?:^|[-_])(STEMS|SRT|MIXES|WEBMIX)(?:[-_]|$)')
STEMS_ONLY = re.compile(r'(?i)(?:^|[-_])STEMS(?:[-_]|$)')
QUAL = re.compile(r'(?i)(?<![a-z0-9])(H264|ProRes)(?![a-z0-9])')
TITLE_EXT = re.compile(r'(?i)\.(?:jpg|jpeg|tif|tiff|png|gif|mp4|mov)$')
SEQUENCE = re.compile(r'(?i)(?:story|sequence|slide|frame)[-_ ]?(\d{1,3})(?!\d)')
RATIO_TOKEN = re.compile(r'(?i)(?:^|[_-])(?:16[-:]9|9[-:]16|1[-:]1|4[-:]5|5[-:]4|4[-:]3|3[-:]4|3[-:]2|2[-:]3|21[-:]9)(?=$|[_-])')
TECH_TOKEN = re.compile(r'(?i)^(?:RGB|sRGB|CMYK|H264|ProRes|Alpha|FullBleed|plate(?:-DE)?|Clean|HQMaster|Master|NoLogo|NOSound|Disclaimer|R128|STEMS|SRT|MIXES|WEBMIX|\d+(?:\.\d+)?s(?:ec(?:onds?)?)?)$')
FILM_TYPE = [
    ('HQMaster', re.compile(r'(?i)HQ[-_]?Master')),
    ('Star-Ending', re.compile(r'(?i)Star[-_]?Ending')),
    ('NoLogo', re.compile(r'(?i)No[-_]?Logo')),
    ('Master', re.compile(r'(?i)(?<!HQ)Master')),
    ('Clean', re.compile(r'(?i)Clean')),
]
VARIATION_PATTERNS = [
    ('HQMaster', re.compile(r'(?i)HQ[-_]?Master')),
    ('Star-Ending', re.compile(r'(?i)Star[-_]?Ending')),
    ('Master', re.compile(r'(?i)(?<!HQ)Master')),
    ('NoLogo', re.compile(r'(?i)No[-_]?Logo')),
    ('Clean', re.compile(r'(?i)Clean')),
    ('NoSound', re.compile(r'(?i)No[-_]?Sound')),
    ('R128', re.compile(r'(?i)(?<![A-Za-z0-9])R128(?![A-Za-z0-9])')),
]


def empty(v):
    return v is None or (isinstance(v, str) and not v.strip())


def unique(values):
    return list(dict.fromkeys(values))


def clean_title(value):
    """Remove only a recognized media extension at the very end."""
    return TITLE_EXT.sub('', str(value)) if not empty(value) else value


def context(title, marker):
    prefix = title[:marker.start()].rstrip('_- ')
    return prefix or 'UNKNOWN'


def numeric_sort(value):
    m = re.fullmatch(r'(\d+)([a-z]?)', str(value), re.I)
    return (int(m.group(1)), ord(m.group(2).lower()) - 96 if m.group(2) else 0) if m else (999999, 0)


def ratio_candidates(title):
    return unique([m.group(1).replace('-', ':') for m in RATIO.finditer(str(title or ''))])


def duration_candidates(title):
    return unique([float(m.group('n')) for m in DURATION.finditer(str(title or ''))])


def rounded_seconds(value):
    try:
        return int(float(value) + 0.5)
    except (TypeError, ValueError):
        return None


def clean_description(title, marker):
    """Return useful title text after a category marker, excluding technical suffixes."""
    tail = title[marker.end():].strip('_- ')
    if not tail:
        return None, 'Description: empty after category marker'
    parts = re.split(r'[_]', tail)
    kept = []
    for part in parts:
        token = part.strip('- ')
        if not token:
            continue
        if RATIO.fullmatch(token.replace('-', ':')) or TECH_TOKEN.fullmatch(token):
            break
        if re.fullmatch(r'(?i)(?:RGB|sRGB|CMYK|DE|EN|IT)', token):
            break
        # Technical descriptors can be embedded after a hyphen (e.g. Clean-NOSound).
        if any(pattern.search(token) for _, pattern in FILM_TYPE) or re.search(r'(?i)^(?:VO-|H264|ProRes|Alpha|plate|FullBleed|NOSound)', token):
            break
        kept.append(token)
    text = ' '.join(kept).replace('-', ' ').strip()
    if not text:
        return None, 'Description: no reliable descriptive text in title'
    return re.sub(r'\s+', ' ', text), None


def source_text(row, keys):
    return ' '.join(str(row.get(key) or '') for key in keys)


def get_duration(row, title, filename, notes):
    """Existing numeric Symphony duration wins; otherwise require one title/filename result."""
    existing = row.get('Duration (seconds)')
    if not empty(existing):
        try:
            number = float(existing)
            return rounded_seconds(number), number, False
        except (TypeError, ValueError):
            notes.append('Film Duration: existing Duration (seconds) is not numeric')
    title_values = duration_candidates(title)
    if len(title_values) == 1:
        return rounded_seconds(title_values[0]), title_values[0], True
    if len(title_values) > 1:
        notes.append('Film Duration: conflicting duration candidates in Title')
        return None, None, False
    file_values = duration_candidates(filename)
    if len(file_values) == 1:
        return rounded_seconds(file_values[0]), file_values[0], True
    if len(file_values) > 1:
        notes.append('Film Duration: conflicting duration candidates in Filename')
        return None, None, False
    notes.append('Film Duration: no recognized duration')
    return None, None, False


def get_quality(row, title, filename, notes):
    for source in (row.get('Quality'), title, filename):
        if not empty(source):
            found = unique([m.group(1).lower() for m in QUAL.finditer(str(source))])
            if len(found) == 1:
                return 'H264' if found[0].lower() == 'h264' else 'ProRes'
            if len(found) > 1:
                notes.append('Quality: conflicting codec markers')
                return None
    notes.append('Quality: no H264 or ProRes marker')
    return None


def get_variation(row, title, filename):
    existing = row.get('Variation')
    if not empty(existing):
        return existing
    source = f'{title} {filename}'
    values = [name for name, pattern in VARIATION_PATTERNS if pattern.search(source)]
    if 'HQMaster' in values and 'Master' in values:
        values.remove('Master')
    for match in re.finditer(r'(?i)(?<![A-Za-z0-9])VO[-_][A-Za-z]{2}(?:[-_][A-Za-z]{2})?(?![A-Za-z])', source):
        values.append(re.sub(r'[_]', '-', match.group(0)).upper())
    return ' / '.join(unique(values)) or None


def identify_master(title, filename):
    source = f'{title} {filename}'
    markers = [
        r'(?i)(?<![A-Za-z0-9])HQ[-_]?Master(?![A-Za-z0-9])',
        r'(?i)(?<![A-Za-z0-9])Master(?![A-Za-z0-9])',
        r'(?i)(?<![A-Za-z0-9])Source(?![A-Za-z0-9])',
        r'(?i)(?<![A-Za-z0-9])Full[-_]?Bleed(?![A-Za-z0-9])',
    ]
    return 'Yes' if any(re.search(pattern, source) for pattern in markers) else None


def extract(row):
    title = str(clean_title(row.get('Title')) or '')
    filename = str(row.get('Filename') or '')
    instr = str(row.get('Instruments') or '')
    out = {key: None for key in NEW}
    notes = []
    expected = []
    category = 'other'
    prefix = ''
    role = ''

    # Keep original category precedence; social categories are additional branches.
    kv = KV.search(title)
    kimi = KIMI.search(title)
    prod = PROD_FILM.search(title)
    tvc = TVC.search(title)
    social_still = SOCIAL_STILL.search(title)
    social_film = SOCIAL_FILM.search(title)
    markers = sum(bool(x) for x in (kv, kimi, prod, tvc, social_still, social_film))
    if markers > 1:
        notes.append('Classification: conflicting title category markers')

    if tvc:
        category, out['PPT Instrument'], prefix = 'tvc', 'TVC', context(title, tvc)
        expected = ['Film Duration', 'Film Ratio', 'Edit Family', 'Asset Role']
    elif prod:
        category, out['PPT Instrument'], prefix = 'film', 'Product Film', context(title, prod)
        expected = ['Film Duration', 'Film Ratio']
    elif social_film:
        category, out['PPT Instrument'], prefix = 'socialfilm', 'Social Film', context(title, social_film)
        expected = ['Film Duration', 'Film Ratio']
    elif social_still:
        category, out['PPT Instrument'], prefix = 'socialstill', 'Social Still', context(title, social_still)
        seq = SEQUENCE.search(title)
        if seq:
            out['Sequence #'] = seq.group(1)
        expected = ['KV Description']
    elif kv:
        category, out['PPT Instrument'], prefix = 'kv', 'Key Visual', context(title, kv)
        out['KV #'] = kv.group(1)
        out['KV Family'] = f'{prefix}|ProdStillUnreal|{out["KV #"]}'
        expected = ['KV #', 'KV Description', 'KV Family']
    elif kimi:
        context_text = source_text(row, ('Title', 'Description', 'Campaign'))
        if re.search(r'Kimi[\s_-]*Antonelli', context_text, re.I):
            category, out['PPT Instrument'], prefix = 'kimi', 'Key Visual', context(title, kimi)
            family = kimi.group(1).capitalize()
            out['KV Family'], out['KV #'] = family, kimi.group(2)
            expected = ['KV Family', 'KV #']
        else:
            notes.append('Kimi classification: family token without confirmed Kimi context')
    elif 'Picture Library' in instr:
        out['PPT Instrument'] = 'Key Visual'

    if category in ('film', 'tvc', 'socialfilm'):
        seconds, source_duration, derived = get_duration(row, title, filename, notes)
        if seconds is not None:
            out['Film Duration'] = f'{seconds}s'
            if derived:
                row['Duration (seconds)'] = seconds
        quality = get_quality(row, title, filename, notes)
        out['Quality'] = quality
        out['Variation'] = get_variation(row, title, filename)
        if category == 'film':
            types = [name for name, pattern in FILM_TYPE if pattern.search(title)]
            if 'HQMaster' in types and 'Master' in types:
                types.remove('Master')
            out['Asset Type'] = ' / '.join(types) or None
        if category == 'tvc':
            support = SUPPORTED.search(title) or SUPPORTED.search(filename)
            role = support.group(1).upper() if support else 'Main'
            out['Asset Role'] = f'Supporting - {role}' if support else 'Main'
            edit_source = f'{title} {filename}'
            no_star = bool(re.search(r'No[-_]?Star', edit_source, re.I))
            no_f1 = bool(re.search(r'no[-_]?F1', edit_source, re.I))
            tactical = bool(re.search(r'Tactical[-_]?Logo[-_]?Ending', edit_source, re.I))
            yt = bool(re.search(r'YT[-_]?Ending', edit_source, re.I))
            if sum((tactical, yt, no_star or no_f1)) > 1:
                notes.append('Edit Family: competing family markers')
            elif tactical:
                out['Edit Family'] = 'Tactical Logo Ending'
            elif yt:
                out['Edit Family'] = 'YT Ending'
            elif no_star and no_f1:
                out['Edit Family'] = 'No-Star / noF1'
            elif no_star:
                out['Edit Family'] = 'No-Star'
            elif no_f1:
                out['Edit Family'] = 'noF1'
            else:
                out['Edit Family'] = 'Standard'
            if support:
                # Supporting packages are retained, but are not expected to carry
                # picture/video ratio, codec, or main-edit metadata.
                expected = ['Asset Role']

    if category in ('kv', 'kimi', 'socialstill', 'film', 'tvc', 'socialfilm'):
        ratios = ratio_candidates(title)
        existing_ratio = row.get('Aspect Ratio')
        if empty(existing_ratio):
            if len(ratios) == 1:
                row['Aspect Ratio'] = ratios[0]
            elif len(ratios) > 1:
                notes.append('Aspect Ratio: multiple candidates ' + ', '.join(ratios))
        ratio = row.get('Aspect Ratio')
        if category in ('film', 'tvc', 'socialfilm') and not (category == 'tvc' and role.startswith('Supporting')):
            if not empty(ratio):
                normalized = str(ratio).replace('-', ':')
                out['Film Ratio'] = normalized
            elif len(ratios) == 1:
                out['Film Ratio'] = ratios[0]
            else:
                notes.append('Film Ratio: no reliable ratio')
        elif category in ('kv', 'kimi', 'socialstill') and empty(ratio):
            notes.append('Aspect Ratio: expected but missing')

    # Prefer real business source columns; never infer them from filenames.
    out['Car Model'] = row.get('Car Model') if not empty(row.get('Car Model')) else (row.get('Car Class') or None)
    out['Motorization'] = row.get('Motorization') if not empty(row.get('Motorization')) else (row.get('Engine Type AMG') or None)
    out['Master'] = row.get('Master') if not empty(row.get('Master')) else identify_master(title, filename)
    ext = Path(filename).suffix.lstrip('.').upper()
    out['Format'] = row.get('Format') if not empty(row.get('Format')) else ('JPG' if ext == 'JPEG' else 'TIF' if ext == 'TIFF' else ext or None)

    # Use explicit source descriptions when they contain meaningful content; otherwise use title text.
    description = row.get('Description')
    description_is_title = not empty(description) and str(description).strip() == title
    description_is_boilerplate = not empty(description) and bool(re.match(r'(?is)^\s*(?:PLEASE NOTE:|This asset is only intended|Taxonomy TBD:)', str(description)))
    desc_value = None if empty(description) or description_is_title or description_is_boilerplate else str(description).strip()
    if category in ('kv', 'kimi', 'socialstill'):
        marker = kv or kimi or social_still
        desc, err = clean_description(title, marker) if marker else (None, None)
        out['KV Description'] = desc_value or desc
        if empty(out['KV Description']) and err:
            notes.append(err)
    elif category in ('film', 'socialfilm'):
        desc, _ = clean_description(title, prod or social_film)
        out['KV Description'] = desc_value or desc
    elif category == 'tvc':
        out['KV Description'] = desc_value

    if category in ('kv', 'socialstill'):
        out['Group Key'] = out['KV Family'] or f'{prefix}|SocialStill'
    elif category == 'kimi':
        out['Group Key'] = f'{prefix}|Kimi|{out["KV Family"]}|{out["KV #"]}'
    elif category == 'film' and out['Film Duration'] and out['Film Ratio']:
        out['Group Key'] = f'{prefix}|ProdFilm|{out["Film Duration"]}|{out["Film Ratio"]}'
    elif category == 'socialfilm' and out['Film Duration'] and out['Film Ratio']:
        out['Group Key'] = f'{prefix}|SocialFilm|{out["Film Duration"]}|{out["Film Ratio"]}'
    elif category == 'tvc' and not role.startswith('Supporting') and out['Film Duration'] and out['Film Ratio'] and out['Edit Family']:
        out['Group Key'] = f'{prefix}|TVC|{out["Film Duration"]}|{out["Edit Family"]}|{out["Film Ratio"]}'
    elif category != 'other' and not (category == 'tvc' and role.startswith('Supporting')):
        notes.append('Group Key: insufficient reliable identifiers')

    if category in ('kv', 'kimi', 'film', 'tvc', 'socialstill', 'socialfilm'):
        if not (category == 'tvc' and role.startswith('Supporting')):
            expected += ['Group Key']
        for field in expected:
            if empty(out.get(field)):
                notes.append(f'{field}: expected but missing')
        if empty(row.get('Hyperlink')):
            notes.append('Symphony Link: expected but missing')
        out['Mapping Status'] = 'Review Required' if notes else 'Mapped'
    else:
        out['Mapping Status'] = 'Review Required' if notes else 'Not Applicable'
    out['Mapping Notes'] = '; '.join(unique(notes)) if notes else None
    return out, category


def parse_family(title, marker_pattern):
    marker = marker_pattern.search(title)
    return context(title, marker) if marker else ''


def link_stems(rows):
    """Assign only an unambiguous STEMS Symphony URL to matching main TVC assets."""
    stems = []
    for rec in rows:
        _, mapped, category, _, work = rec
        title = str(clean_title(work.get('Title')) or '')
        filename = str(work.get('Filename') or '')
        if category == 'tvc' and (STEMS_ONLY.search(title) or STEMS_ONLY.search(filename)):
            duration = mapped.get('Film Duration')
            family = mapped.get('Edit Family')
            prefix = parse_family(title, TVC) or parse_family(filename, TVC)
            stems.append((rec, prefix, duration, family, work.get('Hyperlink')))
    for rec in rows:
        _, mapped, category, _, work = rec
        if category != 'tvc' or str(mapped.get('Asset Role') or '').startswith('Supporting'):
            continue
        title = str(clean_title(work.get('Title')) or '')
        filename = str(work.get('Filename') or '')
        prefix = parse_family(title, TVC) or parse_family(filename, TVC)
        candidates = [url for _, stem_prefix, duration, family, url in stems
                      if stem_prefix == prefix and duration == mapped.get('Film Duration')
                      and family == mapped.get('Edit Family') and not empty(url)]
        candidates = unique(candidates)
        if len(candidates) == 1:
            mapped['Link to Stems'] = candidates[0]
        elif len(candidates) > 1:
            note = 'Link to Stems: multiple matching STEMS assets; left blank'
            mapped['Mapping Notes'] = '; '.join(unique([x for x in [mapped.get('Mapping Notes'), note] if x]))
            mapped['Mapping Status'] = 'Review Required'


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--input', default=DEFAULT_INPUT, type=Path, help=f'Input workbook (default: {DEFAULT_INPUT})')
    p.add_argument('--output', default=DEFAULT_OUTPUT, type=Path, help=f'Output workbook (default: {DEFAULT_OUTPUT})')
    p.add_argument('--sheet', default='Items')
    a = p.parse_args()
    if a.input.resolve() == a.output.resolve():
        p.error('Input and output must be different files')
    if not a.input.is_file():
        p.error(f'Input workbook not found: {a.input}')
    wb = load_workbook(a.input)
    if a.sheet not in wb:
        p.error(f'Sheet {a.sheet!r} not found')
    ws = wb[a.sheet]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    required = ('Asset ID', 'Title', 'Filename', 'Instruments', 'Hyperlink', 'Aspect Ratio', 'Duration (seconds)')
    if any(key not in headers for key in required):
        p.error('Required Symphony columns missing')
    header_col = {name: index + 1 for index, name in enumerate(headers) if name is not None}
    original_columns = len(headers)
    for name in NEW:
        if name not in header_col:
            col = ws.max_column + 1
            cell = ws.cell(1, col, name)
            if original_columns:
                source = ws.cell(1, original_columns)
                if source.has_style:
                    cell._style = copy(source._style)
            else:
                cell.fill = PatternFill('solid', fgColor='243B53')
                cell.font = Font(color='FFFFFF', bold=True)
            ws.column_dimensions[get_column_letter(col)].width = 30 if name not in ('Mapping Notes', 'Group Key') else 65
            header_col[name] = col

    rows = []
    counts = Counter()
    title_cleaned = durations_recovered = qualities_recovered = variations_recovered = 0
    for ri in range(2, ws.max_row + 1):
        original = {key: ws.cell(ri, col).value for key, col in header_col.items()}
        # Include source fields, while keeping any existing enrichment available to helpers.
        for key, col in enumerate(headers, 1):
            if col is not None:
                original[col] = ws.cell(ri, key).value
        working = dict(original)
        cleaned = clean_title(working.get('Title'))
        if cleaned != working.get('Title'):
            title_cleaned += 1
            working['Title'] = cleaned
            ws.cell(ri, header_col['Title']).value = cleaned
        prior_duration = working.get('Duration (seconds)')
        prior_quality = working.get('Quality')
        prior_variation = working.get('Variation')
        mapped, category = extract(working)
        if empty(prior_duration) and not empty(working.get('Duration (seconds)')):
            ws.cell(ri, header_col['Duration (seconds)']).value = working['Duration (seconds)']
            durations_recovered += 1
        if empty(prior_quality) and not empty(mapped.get('Quality')):
            qualities_recovered += 1
        if empty(prior_variation) and not empty(mapped.get('Variation')):
            variations_recovered += 1
        rows.append((ri, mapped, category, original, working))
        counts[category] += 1

    link_stems(rows)

    # Stable, category-aware order values; keep original worksheet rows untouched.
    def order_key(rec):
        ri, mapped, category, orig, work = rec
        ratio = str(work.get('Aspect Ratio') or '')
        master = mapped.get('Master') == 'Yes'
        ratio_rank = {'16:9': 0, '4:5': 1, '9:16': 2, '4:3': 3, '1:1': 4}.get(ratio, 8)
        if category in ('kv', 'kimi'):
            return (mapped.get('Group Key') or '', numeric_sort(mapped.get('KV #')), 0 if master else 1, ratio_rank, ri)
        return (mapped.get('Group Key') or '', 0 if master else 1, ratio_rank, ri)

    grouped = defaultdict(list)
    for record in rows:
        if record[1].get('Group Key'):
            grouped[record[1]['Group Key']].append(record)
    ordered_groups = sorted(grouped, key=lambda group: (next((r[2] for r in grouped[group]), ''),
                                                         numeric_sort(grouped[group][0][1].get('KV #')), group))
    for group_index, key in enumerate(ordered_groups, 1):
        for item_index, record in enumerate(sorted(grouped[key], key=order_key), 1):
            record[1]['Sort Order'] = f'{group_index:06d}.{item_index:04d}'

    for ri, mapped, category, original, working in rows:
        for key in ('Aspect Ratio', 'Duration (seconds)'):
            if empty(original.get(key)) and not empty(working.get(key)):
                ws.cell(ri, header_col[key]).value = working[key]
        for name in NEW:
            cell = ws.cell(ri, header_col[name], mapped.get(name))
            if name in ('Mapping Status', 'Mapping Notes') and mapped.get('Mapping Status') == 'Review Required':
                cell.fill = PatternFill('solid', fgColor='FFF0D9')
            if name == 'Mapping Notes':
                cell.alignment = Alignment(wrap_text=True, vertical='top')

    if ws.auto_filter.ref:
        ws.auto_filter.ref = f'A1:{get_column_letter(ws.max_column)}{ws.max_row}'
    a.output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(a.output)
    status = Counter(record[1]['Mapping Status'] for record in rows)
    stems_count = sum(not empty(record[1].get('Link to Stems')) for record in rows)
    print('Saved:', a.output)
    print('Rows:', len(rows), 'Original columns:', original_columns,
          'Mapped columns:', sum(name in header_col for name in NEW))
    print('Titles cleaned:', title_cleaned)
    print('SocialStill:', counts['socialstill'], 'SocialFilm:', counts['socialfilm'])
    print('Durations recovered:', durations_recovered, 'Quality recovered:', qualities_recovered,
          'Variation recovered:', variations_recovered, 'STEMS links matched:', stems_count)
    print('Categories:', dict(counts))
    print('Mapping status:', dict(status))


if __name__ == '__main__':
    main()
