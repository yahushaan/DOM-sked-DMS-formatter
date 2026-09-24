from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
import pdfplumber
import re, os

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


def norm(s):
    return re.sub(r'\s+', ' ', str(s or '').strip())


def excel_time(s):
    s = norm(s)
    if not s or not TIME_RE.fullmatch(s):
        return None
    n = int(s)
    return n if n <= 2359 else None


def extract_schedule_date(text):
    m = re.search(r'(?i)(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+(\d{2}[./-]\d{2}[./-]\d{4})', text)
    if m:
        return m.group(1).replace('/', '.').replace('-', '.')
    return None


def parse_pdf(pdf_path):
    rows, all_text = [], []
    current_type = current_reg = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=2, y_tolerance=3) or ''
            all_text.append(text)

            for raw in text.splitlines():
                toks = raw.split()
                if not toks:
                    continue

                joined = ' '.join(toks)
                m = re.match(r'^(A320|AT46|AT76|DH8C)\s+Regn\.\s+([A-Z0-9-]+)$', joined, re.I)
                if m:
                    current_type, current_reg = m.group(1).upper(), m.group(2).upper()
                    continue

                pos = next((i for i, tok in enumerate(toks) if FLIGHT_RE.fullmatch(tok)), None)
                if pos is None or current_type is None or current_reg is None:
                    continue

                rest = toks[pos + 1:]
                if len(rest) < 6 or len(rest) < 3 or not AIRPORT_RE.fullmatch(rest[2]):
                    continue

                frm = rest[2].upper()
                k, times = 3, []
                while k < len(rest) and TIME_RE.fullmatch(rest[k]):
                    times.append(rest[k])
                    k += 1

                if len(times) not in (2, 4) or k >= len(rest) or not AIRPORT_RE.fullmatch(rest[k]):
                    continue

                rows.append({
                    'flight': toks[pos].upper(),
                    'frm': frm,
                    'std': times[0],
                    'blkoff': times[1] if len(times) == 4 else '',
                    'blkin': times[2] if len(times) == 4 else '',
                    'sta': times[-1],
                    'to': rest[k].upper(),
                    'type': current_type,
                    'reg': current_reg,
                })

    if not rows:
        raise ValueError('No flight rows were found. This application is tuned to the iFlight Lite Daily OPS Schedule PDF format.')
    return extract_schedule_date('\n'.join(all_text)), rows


def normalize_flight_number(flight):
    return flight.upper().replace(' ', '')


def combine_flights(legs):
    nums = []
    for leg in legs:
        n = normalize_flight_number(leg['flight'])
        if n not in nums:
            nums.append(n)
    if not nums:
        return ''
    if len(nums) == 1:
        return nums[0]
    if len(nums) == 2 and len(nums[0]) == len(nums[1]) and nums[0][:-1] == nums[1][:-1]:
        return nums[0] + '/' + nums[1][-1]
    return '/'.join(nums)


def display_aircraft_type(ac_type):
    return {'AT46': 'AT-46', 'AT76': 'AT-76', 'DH8C': 'DH8', 'A320': 'A320'}.get(ac_type, ac_type)


def display_registration(reg):
    reg = reg.upper()
    if reg.startswith('8Q-'):
        return reg
    if reg.startswith('8Q'):
        return '8Q-' + reg[2:]
    return reg


def trim_inbound_group(legs):
    # For an inbound block, ignore unrelated earlier flying and retain only
    # the consecutive sectors belonging to the final MLE-bound flight number.
    if not legs or legs[0]['frm'] == 'MLE' or legs[-1]['to'] != 'MLE':
        return legs
    final_flight = legs[-1]['flight']
    start = len(legs) - 1
    while start > 0 and legs[start - 1]['flight'] == final_flight:
        start -= 1
    return legs[start:]


def group_rotations(rows, include_a320=False):
    if not include_a320:
        rows = [r for r in rows if r['type'] != 'A320']

    by_reg, reg_order = {}, []
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
                groups.append(trim_inbound_group(current))
                current = []
        # Keep a final departure from MLE when its return is outside this schedule.
        if current and current[0]['frm'] == 'MLE':
            groups.append(current)

    result = []
    for legs in groups:
        if not legs:
            continue

        route = []
        for leg in legs:
            for airport in (leg['frm'], leg['to']):
                if airport and (not route or route[-1] != airport):
                    route.append(airport)

        first, last = legs[0], legs[-1]
        from_mle, to_mle = first['frm'] == 'MLE', last['to'] == 'MLE'

        result.append({
            'flight': combine_flights(legs),
            'type': display_aircraft_type(first['type']),
            'reg': display_registration(first['reg']),
            'routing': ' '.join(route),
            'std': excel_time(first['std']) if from_mle else None,
            'atd': excel_time(first['blkoff']) if from_mle and first['blkoff'] else None,
            'sta': excel_time(last['sta']) if to_mle else None,
            'eta': excel_time(last['blkin']) if to_mle and last['blkin'] else None,
        })

    type_priority = {'AT-76': 0, 'AT-46': 1, 'DH8': 2, 'A320': 3}
    result.sort(key=lambda x:
                (0, x['sta'], type_priority.get(x['type'], 9), x['flight'])
                if x['std'] is None else
                (1, x['std'], type_priority.get(x['type'], 9), x['flight']))
    return result


def title_for_date(date_string):
    if not date_string:
        return 'SCHEDULE'
    dt = datetime.strptime(date_string, '%d.%m.%Y')
    d = dt.day
    suffix = 'TH' if 10 <= d % 100 <= 20 else {1: 'ST', 2: 'ND', 3: 'RD'}.get(d % 10, 'TH')
    return f'SCHEDULE {d}{suffix}'


def make_excel(blocks, date_string):
    wb = load_workbook(TEMPLATE)
    ws = wb.active
    ws.title = 'Schedule'
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
    result = error = preview = None
    if request.method == 'POST':
        f = request.files.get('pdf')
        include_a320 = request.form.get('include_a320') == '1'
        if not f or not f.filename.lower().endswith('.pdf'):
            error = 'Please select a PDF schedule.'
        else:
            pdf_path = UPLOADS / secure_filename(f.filename)
            f.save(pdf_path)
            try:
                date_string, rows = parse_pdf(pdf_path)
                blocks = group_rotations(rows, include_a320=include_a320)
                if not blocks:
                    raise ValueError('No schedule blocks remained after applying the selected rules.')
                out = make_excel(blocks, date_string)
                preview = blocks
                result = {
                    'filename': out.name, 'count': len(blocks), 'rows': len(rows),
                    'date': date_string or 'unknown', 'a320': include_a320,
                }
            except Exception as exc:
                error = str(exc)
    return render_template('index.html', result=result, error=error, preview=preview)


@app.route('/download/<path:name>')
def download(name):
    p = OUT / secure_filename(name)
    if not p.exists():
        return 'File not found', 404
    return send_file(p, as_attachment=True, download_name=p.name)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', '8080')), debug=False)
