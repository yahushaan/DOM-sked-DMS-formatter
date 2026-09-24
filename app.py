from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
import pdfplumber
import re, os, shutil

BASE = Path(__file__).resolve().parent
TEMPLATE = BASE / 'templates' / 'MDV_Schedule_Template.xlsx'
OUT = BASE / 'output'
UPLOADS = BASE / 'uploads'
OUT.mkdir(exist_ok=True)
UPLOADS.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'mdv-internal-schedule-change-me')
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024

FLIGHT_RE = re.compile(r'^Q\d{3,5}$', re.I)
TIME_RE = re.compile(r'^\d{3,4}$')
AIRPORT_RE = re.compile(r'^[A-Z]{3}$')
AIRCRAFT_TYPES = {'A320', 'AT46', 'AT76', 'DH8C'}


def norm(s):
    return re.sub(r'\s+', ' ', str(s or '').strip())


def excel_time(s):
    s = norm(s)
    if not s or not TIME_RE.fullmatch(s):
        return None
    n = int(s)
    if n > 2359:
        return None
    return n


def extract_schedule_date(text):
    # Prefer the operating/schedule date, not the printed date.
    m = re.search(r'(?i)(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+(\d{2}[./-]\d{2}[./-]\d{4})', text)
    if m:
        return m.group(1).replace('/', '.').replace('-', '.')
    m = re.search(r'(?i)FOR\s+\d{1,2}\s+[A-Z]{3}', text)
    if m:
        # No year in this form; leave date unknown rather than using the printed timestamp.
        return None
    return None


def parse_pdf(pdf_path):
    """Parse the iFlight Lite Daily OPS Schedule format used by the sample PDF.

    The PDF has aircraft section headers such as 'AT46 Regn. 8QIAV'.
    Each flight line has: Flt nbr, Carrier, T, Frm, STD, [BlkOff, BlkIn], STA, To, Blk.
    When Block Off/In are blank, the text extractor collapses those empty columns;
    therefore the parser determines whether there are 2 or 4 time fields before To.
    """
    rows = []
    current_type = None
    current_reg = None
    all_text = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=2, y_tolerance=3) or ''
            all_text.append(text)
            if not extract_schedule_date('\n'.join(all_text)):
                pass

            for raw in text.splitlines():
                toks = raw.split()
                if not toks:
                    continue

                # Aircraft section header.
                joined = ' '.join(toks)
                m = re.match(r'^(A320|AT46|AT76|DH8C)\s+Regn\.\s+([A-Z0-9-]+)$', joined, re.I)
                if m:
                    current_type = m.group(1).upper()
                    current_reg = m.group(2).upper()
                    continue

                # Find the flight number anywhere in the line. This handles the first
                # row of a section, where the date and registration may precede it.
                pos = next((i for i, tok in enumerate(toks) if FLIGHT_RE.fullmatch(tok)), None)
                if pos is None or current_type is None or current_reg is None:
                    continue

                rest = toks[pos + 1:]
                # Expected: Carrier, T, Frm, one/two/four time fields, To, Block duration...
                if len(rest) < 6:
                    continue
                if len(rest) < 3 or not AIRPORT_RE.fullmatch(rest[2]):
                    continue
                frm = rest[2].upper()

                k = 3
                times = []
                while k < len(rest) and TIME_RE.fullmatch(rest[k]):
                    times.append(rest[k])
                    k += 1
                if len(times) not in (2, 4):
                    continue
                if k >= len(rest) or not AIRPORT_RE.fullmatch(rest[k]):
                    continue
                to = rest[k].upper()

                row = {
                    'flight': toks[pos].upper(),
                    'frm': frm,
                    'std': times[0],
                    'blkoff': times[1] if len(times) == 4 else '',
                    'blkin': times[2] if len(times) == 4 else '',
                    'sta': times[-1],
                    'to': to,
                    'type': current_type,
                    'reg': current_reg,
                }
                rows.append(row)

    full_text = '\n'.join(all_text)
    date_string = extract_schedule_date(full_text)
    if not rows:
        raise ValueError('No flight rows were found. This application is tuned to the iFlight Lite Daily OPS Schedule PDF format.')
    return date_string, rows


def normalize_flight_number(flight):
    # The required schedule uses Q2 as the carrier prefix and the remaining digits.
    # Q2706 -> Q2 706; Q2220 -> Q2 220; Q22302 -> Q2 2302.
    digits = re.sub(r'^Q', '', flight.upper())
    if digits.startswith('2'):
        digits = digits[1:]
    return digits


def combine_flights(legs):
    nums = []
    for leg in legs:
        n = normalize_flight_number(leg['flight'])
        if n not in nums:
            nums.append(n)
    if not nums:
        return ''
    if len(nums) == 1:
        return 'Q2 ' + nums[0]
    # Required style for paired sectors: Q2 706/7, Q2 2232/3, Q2 2448/9.
    if len(nums) == 2 and len(nums[0]) == len(nums[1]) and nums[0][:-1] == nums[1][:-1]:
        return 'Q2 ' + nums[0] + '/' + nums[1][-1]
    return 'Q2 ' + '/'.join(nums)


def group_rotations(rows, include_a320=False):
    """Create schedule blocks by aircraft rotation.

    A block starts with an arrival or departure and ends when the aircraft reaches MLE.
    This is what converts consecutive sectors into one formatted schedule line.
    The sample's required Schedule excludes A320/international rotations, so A320 is
    excluded by default; the UI provides an option to include it.
    """
    if not include_a320:
        rows = [r for r in rows if r['type'] != 'A320']

    # Preserve the PDF's aircraft/row order, then group within each registration.
    by_reg = {}
    reg_order = []
    for r in rows:
        if r['reg'] not in by_reg:
            by_reg[r['reg']] = []
            reg_order.append(r['reg'])
        by_reg[r['reg']].append(r)

    groups = []
    for reg in reg_order:
        current = []
        for r in by_reg[reg]:
            current.append(r)
            if r['to'] == 'MLE':
                groups.append(current)
                current = []
        if current:
            groups.append(current)

    result = []
    for legs in groups:
        route = []
        for leg in legs:
            for airport in (leg['frm'], leg['to']):
                if airport and (not route or route[-1] != airport):
                    route.append(airport)
        first, last = legs[0], legs[-1]
        from_mle = first['frm'] == 'MLE'
        to_mle = last['to'] == 'MLE'
        result.append({
            'flight': combine_flights(legs),
            'type': first['type'],
            'reg': re.sub(r'^8Q', '', first['reg'], flags=re.I),
            'routing': ' '.join(route),
            'std': excel_time(first['std']) if from_mle else None,
            'atd': excel_time(first['blkoff']) if from_mle and first['blkoff'] else None,
            'sta': excel_time(last['sta']) if to_mle else None,
            'eta': excel_time(last['blkin']) if to_mle and last['blkin'] else None,
        })

    # Required format: arrivals first by STA, then departures by STD.
    # For an identical departure time, the supplied 23 SEP template places AT76
    # before AT46; retain that tie-break so the generated sheet matches the sample.
    type_priority = {'AT76': 0, 'AT46': 1, 'DH8C': 2, 'A320': 3}
    result.sort(key=lambda x: (0, x['sta'], type_priority.get(x['type'], 9), x['flight'])
               if x['std'] is None else
               (1, x['std'], type_priority.get(x['type'], 9), x['flight']))
    return result


def title_for_date(date_string):
    if not date_string:
        return 'SCHEDULE'
    dt = datetime.strptime(date_string, '%d.%m.%Y')
    d = dt.day
    if 10 <= d % 100 <= 20:
        suffix = 'TH'
    else:
        suffix = {1: 'ST', 2: 'ND', 3: 'RD'}.get(d % 10, 'TH')
    return f'SCHEDULE {d}{suffix}'


def make_excel(blocks, date_string):
    wb = load_workbook(TEMPLATE)
    ws = wb.active
    ws.title = 'Schedule'

    # The template is intentionally kept intact; only the schedule body/title are replaced.
    ws['D3'] = title_for_date(date_string)
    for row in range(6, max(ws.max_row, 155) + 1):
        for col in range(3, 11):
            ws.cell(row=row, column=col).value = None

    for i, b in enumerate(blocks, start=6):
        for j, value in enumerate([
            b['flight'], b['type'], b['reg'], b['routing'],
            b['std'], b['atd'], b['sta'], b['eta']
        ], start=3):
            ws.cell(i, j).value = value

    end_row = max(5, 5 + len(blocks))
    ws.print_area = f'C3:J{end_row}'
    ws.freeze_panes = 'C6'

    date_tag = date_string.replace('.', '_') if date_string else datetime.now().strftime('%d_%m_%Y')
    out = OUT / f'MDV_SKED_{date_tag}.xlsx'
    wb.save(out)
    return out


@app.route('/', methods=['GET', 'POST'])
def index():
    result = None
    error = None
    preview = None
    if request.method == 'POST':
        f = request.files.get('pdf')
        include_a320 = request.form.get('include_a320') == '1'
        if not f or not f.filename.lower().endswith('.pdf'):
            error = 'Please select a PDF schedule.'
        else:
            name = secure_filename(f.filename)
            pdf_path = UPLOADS / name
            f.save(pdf_path)
            try:
                date_string, rows = parse_pdf(pdf_path)
                blocks = group_rotations(rows, include_a320=include_a320)
                if not blocks:
                    raise ValueError('No schedule blocks remained after applying the selected rules.')
                out = make_excel(blocks, date_string)
                preview = blocks
                result = {
                    'filename': out.name,
                    'count': len(blocks),
                    'rows': len(rows),
                    'date': date_string or 'unknown',
                    'a320': include_a320,
                }
            except Exception as exc:
                error = str(exc)
    return render_template('index.html', result=result, error=error, preview=preview)


@app.route('/download/<path:name>')
def download(name):
    safe = secure_filename(name)
    p = OUT / safe
    if not p.exists():
        return 'File not found', 404
    return send_file(p, as_attachment=True, download_name=p.name)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', '8080')), debug=False)
