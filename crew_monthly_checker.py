# ==========================================
# crew_monthly_checker.py
# 버전: v1.1 (2026-07-14)  — crew_check.js v17 과 룰 동기화
# 기능: CMS 편조점검 월간 자동 조회 (규정위반/내부위반 엑셀 저장)
# 문의: 승무계획팀
# ==========================================
import asyncio
import re
import calendar
from datetime import datetime
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright, TimeoutError as PWTimeout
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from tkinter import messagebox
import tkinter as tk
import os

CMS_URL  = "https://crew.eastarjet.com/cms/Admin/Schedule/CrewPairs/CrewPairList.php"
HEADLESS = False

# ==========================================
# 편조점검 설정 (crew_check.js 와 동일)
# ==========================================
CFG = {
    "A": {"YNT","DSN","DAT","CGO","TXN","CGQ","SHE","HRB","MDC","KOJ","KMJ","IZO","TKS","TAE","CXR","DYG","DLC","YNJ","HKG","BSZ","ALA","MFM"},
    "B": {"NTG","NRT","OKA","TSA","DAD","FUK","AOJ","PUS"},
    "C": {"PVG","KIX","CTS","KUV","ICN","GMP","CJJ","BKK","CNX","TPE","PQC","CJU"},
    "cxrBan" : {"신윤식","정진우"},
    "dadBan" : {"장준욱"},
    "foAonly": {"김상겸"},
    "foABonly":{"신영근"},
    "qa"     : {"박지현","신현욱","박승훈","신준서"},
    "cp"     : {"황종식","성기중","이재환","이태우"},
    # 세이프티 불가 명단 (2026-07-31 기준)
    "spBan"  : {"김창중","이주화","양병모","엄태국","김우영","최은총","장재봉","이창민",
                "이한솔","정종성","김공주","김총화","김재영","이웅배","김민재","한다영","최도현"},
    # 세이프티 예외자 (불가 명단이지만 세이프티 가능)
    "spOK"   : {"엄태국","양병모"},
    "gradeOverride": {},
}

NAME_RE  = re.compile(r'[가-힣]{2,5}(?:[ABCX](?:LV|ALV|CLV)?)?')
GRADE_RE = re.compile(r'^[가-힣]{2,5}([ABCX])(?:LV|ALV|CLV)?')

def get_name(s):
    return re.sub(r'[ABCX](?:LV|ALV|CLV)?.*$', '', s)

def get_site_grade(s):
    m = GRADE_RE.match(s)
    return m.group(1) if m else ''

def get_grade(s):
    n = get_name(s)
    if n in CFG["gradeOverride"]:
        return CFG["gradeOverride"][n]
    return get_site_grade(s)

def is_junk(s):
    return bool(re.match(r'^\d{1,2}/\d{1,2}', s)) or '편조' in s or '점검' in s or len(s) == 0

def parse_table(html):
    soup = BeautifulSoup(html, 'html.parser')
    tables = soup.find_all('table')
    if len(tables) < 2:
        return []

    rows = []
    for tr in tables[1].find_all('tr'):
        cells = tr.find_all('td')
        if not cells or len(cells) < 3:
            t = tr.get_text(' ', strip=True)
            if t:
                rows.append(t)
            continue

        def cell_txt(c):
            txt = re.sub(r' {2,}', ' ', c.get_text(' ', strip=True))
            # CMS HTML: 이름과 등급 사이 공백 제거 ("정병국 A" → "정병국A")
            txt = re.sub(r'([가-힣]) ([ABCX](?:LV|ALV|CLV)?)\b', r'\1\2', txt)
            return txt

        cap_cell   = cell_txt(cells[0])
        fo_cell    = cell_txt(cells[1])
        extra_cell = cell_txt(cells[2]) if len(cells) > 2 else ''

        caps   = NAME_RE.findall(cap_cell)
        fos    = NAME_RE.findall(fo_cell)
        extras = NAME_RE.findall(extra_cell)

        rest = ' '.join(cell_txt(cells[i]) for i in range(3, len(cells)))
        ordered = []
        for k in range(max(len(caps), len(fos))):
            if k < len(caps): ordered.append(caps[k])
            if k < len(fos):  ordered.append(fos[k])
        ordered += extras
        line = (' '.join(ordered) + ' ' + rest).strip()
        if line:
            rows.append(line)
    return rows

TOKEN_RE = re.compile(
    r'(\d{2}:\d{2})'
    r'|([A-Z]{3,4}/[A-Z]{3,4})'
    r'|(\d{3,4})(?![\d:])'
    r'|([가-힣]{2,5}(?:[ABCX](?:LV|ALV|CLV)?)?)'
)

def parse_blocks(rows):
    merged = []
    for L in rows:
        if re.match(r'^(?:LV|ALV|CLV)$', L) and merged and re.search(r'[가-힣]{2,5}[ABCX]?$', merged[-1]):
            merged[-1] += L
        else:
            merged.append(L)
    clean = [L for L in merged if not is_junk(L)]

    typed = []
    for line in clean:
        for m in TOKEN_RE.finditer(line):
            if m.group(1):   typed.append(('time',   m.group(1)))
            elif m.group(2): typed.append(('route',  m.group(2)))
            elif m.group(3): typed.append(('flight', m.group(3)))
            elif m.group(4): typed.append(('name',   m.group(4)))

    blocks, i, N = [], 0, len(typed)
    while i < N:
        ns = []
        while i < N and typed[i][0] == 'name':
            ns.append(typed[i][1]); i += 1
        if not ns:
            fl0 = []
            while i < N and typed[i][0] == 'flight':
                f0 = typed[i][1]; i += 1
                if i < N and typed[i][0] == 'route':
                    r0 = typed[i][1]; i += 1
                    while i < N and typed[i][0] == 'time': i += 1
                    fl0.append({'fl': f0, 'rt': r0})
            if fl0 and blocks:
                blocks[-1]['flights'].extend(fl0)
            while i < N and typed[i][0] not in ('name', 'flight'): i += 1
            continue
        fl1 = []
        while i < N and typed[i][0] == 'flight':
            f1 = typed[i][1]; i += 1
            if i < N and typed[i][0] == 'route':
                r1 = typed[i][1]; i += 1
                while i < N and typed[i][0] == 'time': i += 1
                fl1.append({'fl': f1, 'rt': r1})
        blocks.append({'names': ns, 'flights': fl1})

    mains, solos = [], []
    for b in blocks:
        if len(b['names']) >= 2:
            mains.append({'cap': b['names'][0], 'fo': b['names'][1],
                          'extra': b['names'][2:], 'flights': b['flights']})
        elif len(b['names']) == 1 and b['flights']:
            solos.append(b)

    for s in solos:
        sfl  = {f['fl'] for f in s['flights']}
        best, bsc = None, -1
        for m in mains:
            mfl = {f['fl'] for f in m['flights']}
            cnt = len(sfl & mfl)
            if cnt == len(sfl) and cnt > bsc:
                best, bsc = m, cnt
        all_xor = all(get_grade(n) in ('X', '') for n in s['names'])
        if all_xor:
            s['asSolo'] = True
        elif best:
            best['extra'].extend(s['names'])
        else:
            s['asSolo'] = True

    result = []
    for m in mains:
        all_n   = [m['cap'], m['fo']] + m['extra']
        graded  = [n for n in all_n if get_grade(n) in ('A','B','C')]
        nograde = [n for n in all_n if get_grade(n) not in ('A','B','C','X')]
        if len(graded) >= 4:
            result.append({'cap': graded[0], 'fo': graded[1], 'extra': nograde, 'flights': m['flights']})
            result.append({'cap': graded[2], 'fo': graded[3], 'extra': [],      'flights': m['flights']})
        else:
            result.append(m)
    for s in solos:
        if s.get('asSolo'):
            result.append({'isSolo': True, 'names': s['names'], 'flights': s['flights']})
    return result


def check(blocks):
    violations, internalV = [], []
    seen = {'cc': set(), 'cf': set(), 'aa': set(), 'io': set()}
    fl_set = set()

    for b in blocks:
        if b.get('isSolo'):
            fl_set.update(f['fl'] for f in b['flights'])
            continue

        cap_n, cap_g = get_name(b['cap']), get_grade(b['cap'])
        fo_n,  fo_g  = get_name(b['fo']),  get_grade(b['fo'])
        fo_eff = 'SKIP' if fo_g == 'X' else (fo_g if fo_g else '')
        pair   = f"{b['cap']}/{b['fo']}"
        fls    = '/'.join(f['fl'] for f in b['flights'])

        # ── 세이프티 체크 (등급 스킵보다 먼저 수행) ──
        # 기장/부기장/기타 어느 자리든 무등급(훈련생) 또는 X등급(관숙/DH)이 있으면 감지
        has_trainee = (
            cap_g in ('', 'X')
            or fo_g in ('', 'X')
            or any(get_grade(e) in ('', 'X') for e in b.get('extra', []))
        )
        if has_trainee:
            for raw in (b['cap'], b['fo']):
                nm, g = get_name(raw), get_grade(raw)
                if g in ('', 'X'):          # 본인이 훈련생이면 스킵
                    continue
                if nm not in CFG['spBan']:
                    continue
                if nm in CFG['spOK']:
                    internalV.append({'type':'참고','note':True,
                                      'detail':'SP 예외자 (세이프티 가능 - 조치 불필요)',
                                      'fl':fls,'pair':f"{nm} / {pair}"})
                else:
                    internalV.append({'type':'내부위반',
                                      'detail':'세이프티 불가자 + 훈련/관숙 동승 (확인 필요)',
                                      'fl':fls,'pair':f"{nm} / {pair}"})

        # 기장 또는 부기장 등급 없으면 훈련생 페어링 → 규정위반 체크 스킵
        if cap_g == '' or (fo_g == '' and fo_eff == ''):
            continue

        for flt in b['flights']:
            fl_set.add(flt['fl'])
            org, dst = flt['rt'].split('/') if '/' in flt['rt'] else (flt['rt'], '')

            if cap_g == 'C':
                ok = fo_eff == 'A'
                if not ok and fo_eff in ('SKIP', ''):
                    obs = next((e for e in b.get('extra',[]) if get_grade(e) == 'A'), None)
                    if obs: ok = True
                ck = f"cc|{pair}"
                if ck not in seen['cc']:
                    seen['cc'].add(ck)
                if not ok:
                    msg = 'C기장 관숙 편성 위반(FO A 동승 필요)' if fo_eff in ('SKIP','') else 'C기장 페어링 위반'
                    violations.append({'type':'규정위반','detail':msg,'fl':flt['fl'],'pair':pair,'ap':''})
                for ap in (org, dst):
                    if ap in CFG['B']:
                        violations.append({'type':'규정위반','detail':'B공항 C기장 위반','fl':flt['fl'],'pair':pair,'ap':ap})
                    if ap and ap not in CFG['A'] and ap not in CFG['B'] and ap not in CFG['C']:
                        violations.append({'type':'규정위반','detail':'C기장 분류외 공항 위반','fl':flt['fl'],'pair':pair,'ap':ap})

            if fo_eff == 'C':
                ok2 = cap_g == 'A'
                ck2 = f"cf|{pair}"
                if ck2 not in seen['cf']:
                    seen['cf'].add(ck2)
                if not ok2:
                    violations.append({'type':'규정위반','detail':'C부기장 페어링 위반','fl':flt['fl'],'pair':pair,'ap':''})
                for ap in (org, dst):
                    if ap and ap not in CFG['A'] and ap not in CFG['B'] and ap not in CFG['C']:
                        violations.append({'type':'규정위반','detail':'C부기장 분류외 공항 위반','fl':flt['fl'],'pair':pair,'ap':ap})

            for ap in (org, dst):
                if ap in CFG['A']:
                    cok = cap_g == 'A'
                    fok = fo_eff in ('A', 'SKIP')
                    k = f"aa|{pair}|{ap}"
                    if k not in seen['aa']:
                        seen['aa'].add(k)
                    if not cok:
                        violations.append({'type':'규정위반','detail':'A공항 기장 등급 위반','fl':flt['fl'],'pair':pair,'ap':ap})
                    if not fok:
                        violations.append({'type':'규정위반','detail':'A공항 부기장 등급 위반','fl':flt['fl'],'pair':pair,'ap':ap})
                if ap == 'CXR' and cap_n in CFG['cxrBan']:
                    internalV.append({'type':'내부위반','detail':'CXR 금지','fl':flt['fl'],'pair':b['cap'],'ap':''})
                if ap == 'DAD' and cap_n in CFG['dadBan']:
                    internalV.append({'type':'내부위반','detail':'DAD 금지','fl':flt['fl'],'pair':b['cap'],'ap':''})

            if cap_n in CFG['foAonly']:
                if fo_eff != 'A':
                    internalV.append({'type':'내부위반','detail':f'{cap_n} FO제한위반','fl':flt['fl'],'pair':b['fo'],'ap':''})
            if cap_n in CFG['foABonly']:
                if fo_eff not in ('A','B'):
                    internalV.append({'type':'내부위반','detail':f'{cap_n} FO제한위반','fl':flt['fl'],'pair':b['fo'],'ap':''})

    return violations, internalV


def save_excel(all_results, year, month):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{year}년{month:02d}월_위반사항"

    hdr_font  = Font(name='맑은 고딕', bold=True, color='FFFFFF', size=10)
    hdr_fill  = PatternFill('solid', start_color='1F3864')
    v_fill    = PatternFill('solid', start_color='3A1E1E')
    i_fill    = PatternFill('solid', start_color='3A2E1E')
    n_fill    = PatternFill('solid', start_color='1E2A3A')
    center    = Alignment(horizontal='center', vertical='center')
    wrap      = Alignment(wrap_text=True, vertical='center')

    headers = ['날짜', '구분', '편명', '위반유형', '페어링', '공항']
    widths  = [12,     8,     12,     32,         28,       8]
    for col, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(1, col, h)
        c.font, c.fill, c.alignment = hdr_font, hdr_fill, center
        ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[1].height = 20

    row_idx = 2
    total_v = total_i = 0

    for date_str, violations, internalV in all_results:
        for v in violations:
            fill = v_fill
            ws.cell(row_idx, 1, date_str).alignment = center
            ws.cell(row_idx, 2, '🚨규정위반').alignment = center
            ws.cell(row_idx, 3, v['fl']).alignment = center
            ws.cell(row_idx, 4, v['detail']).alignment = wrap
            ws.cell(row_idx, 5, v['pair']).alignment = wrap
            ws.cell(row_idx, 6, v.get('ap','')).alignment = center
            for col in range(1, 7):
                ws.cell(row_idx, col).fill = fill
                ws.cell(row_idx, col).font = Font(name='맑은 고딕', size=10)
            ws.row_dimensions[row_idx].height = 16
            row_idx += 1
            total_v += 1

        for v in internalV:
            is_note = v.get('note', False)
            fill = n_fill if is_note else i_fill
            ws.cell(row_idx, 1, date_str).alignment = center
            ws.cell(row_idx, 2, 'ℹ️참고' if is_note else '⚠️내부위반').alignment = center
            ws.cell(row_idx, 3, v['fl']).alignment = center
            ws.cell(row_idx, 4, v['detail']).alignment = wrap
            ws.cell(row_idx, 5, v['pair']).alignment = wrap
            ws.cell(row_idx, 6, v.get('ap','')).alignment = center
            for col in range(1, 7):
                ws.cell(row_idx, col).fill = fill
                ws.cell(row_idx, col).font = Font(name='맑은 고딕', size=10)
            ws.row_dimensions[row_idx].height = 16
            row_idx += 1
            if not is_note:
                total_i += 1

    if row_idx == 2:
        ws.cell(2, 1, '✅ 위반사항 없음')
        ws.cell(2, 1).font = Font(name='맑은 고딕', bold=True, color='4ADE80', size=11)

    # 요약 시트
    ws2 = wb.create_sheet('요약')
    ws2['A1'] = f"{year}년 {month:02d}월 편조점검 결과"
    ws2['A1'].font = Font(name='맑은 고딕', bold=True, size=13)
    ws2['A3'] = '조회 일수'
    ws2['B3'] = len(all_results)
    ws2['A4'] = '규정위반 건수'
    ws2['B4'] = total_v
    ws2['B4'].font = Font(color='FF6B6B' if total_v else '4ADE80', bold=True)
    ws2['A5'] = '내부위반 건수'
    ws2['B5'] = total_i
    ws2['B5'].font = Font(color='FFD166' if total_i else '4ADE80', bold=True)
    ws2['A6'] = '생성일시'
    ws2['B6'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    ws2.column_dimensions['A'].width = 18
    ws2.column_dimensions['B'].width = 20
    default_font = Font(name='맑은 고딕', size=11)
    for row in ws2.iter_rows(min_row=3, max_row=6):
        for cell in row:
            if not cell.font or not cell.font.name:
                cell.font = default_font

    out_path = os.path.join(
        os.path.expanduser('~'), 'Desktop',
        f"편조점검_{year}{month:02d}.xlsx"
    )
    wb.save(out_path)
    return out_path, total_v, total_i


def get_target_month():
    """이번 달 / 다음 달 선택 팝업"""
    root = tk.Tk()
    root.withdraw()
    today = datetime.now()
    this_y, this_m = today.year, today.month
    if this_m == 12:
        next_y, next_m = this_y + 1, 1
    else:
        next_y, next_m = this_y, this_m + 1

    answer = messagebox.askquestion(
        "조회 월 선택",
        f"조회할 월을 선택하세요.\n\n"
        f"  [예]    다음 달 ({next_y}년 {next_m:02d}월, 1일~말일)\n"
        f"  [아니오] 이번 달 ({this_y}년 {this_m:02d}월, 오늘~말일)",
        icon="question"
    )
    root.destroy()
    if answer == "yes":
        last = calendar.monthrange(next_y, next_m)[1]
        return next_y, next_m, 1, last
    else:
        last = calendar.monthrange(this_y, this_m)[1]
        return this_y, this_m, today.day, last


async def main():
    print('='*50)
    print('✈  편조점검 월간 자동 조회 v1.0')
    print('   (2026-06-24) | 문의: 승무계획팀')
    print('='*50)

    today = datetime.now()
    year, month, start_day, last_day = get_target_month()
    print(f"\n조회 대상: {year}년 {month:02d}월 ({start_day}일 ~ {last_day}일)\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=HEADLESS, args=[
            '--disable-blink-features=AutomationControlled',
            '--no-sandbox', '--window-size=1280,800'
        ])
        context = await browser.new_context(
            locale='ko-KR', timezone_id='Asia/Seoul',
            viewport={'width': 1280, 'height': 800}
        )
        page = await context.new_page()

        # 로그인 대기
        await page.goto(CMS_URL, wait_until='domcontentloaded', timeout=20000)
        print('브라우저가 열렸습니다.')
        print('CMS에 로그인 후 엔터를 눌러주세요...')
        await asyncio.get_event_loop().run_in_executor(None, input, '  [로그인 완료 후 엔터] ')

        all_results = []

        for day in range(start_day, last_day + 1):
            date_str = f"{year}/{month:02d}/{day:02d}"
            url = f"{CMS_URL}?d={year}-{month:02d}-{day:02d}&s=true&e=true&a=true"

            print(f"[{day:02d}/{last_day}] {date_str} 조회 중...", end=' ', flush=True)

            try:
                await page.goto(url, wait_until='networkidle', timeout=30000)
                await page.wait_for_timeout(2500)

                # 세션 만료 체크
                body = await page.inner_text('body')
                if 'SESSION EXPIRED' in body or 'logout' in page.url:
                    print('⚠️  세션 만료 → 재로그인 후 엔터')
                    await asyncio.get_event_loop().run_in_executor(None, input, '  [재로그인 후 엔터] ')
                    await page.goto(url, wait_until='networkidle', timeout=30000)
                    await page.wait_for_timeout(2500)

                html = await page.content()
                rows = parse_table(html)

                if not rows:
                    print('데이터 없음 (스킵)')
                    continue

                blocks = parse_blocks(rows)
                violations, internalV = check(blocks)

                if violations or internalV:
                    all_results.append((date_str, violations, internalV))
                    n_int = sum(1 for v in internalV if not v.get('note'))
                    n_note = len(internalV) - n_int
                    note_txt = f" / ℹ️ 참고 {n_note}건" if n_note else ''
                    print(f"🚨 규정위반 {len(violations)}건 / ⚠️ 내부위반 {n_int}건{note_txt}")
                    for v in violations:
                        ap = f" ({v['ap']})" if v.get('ap') else ''
                        print(f"    🚨 [{v['fl']}] {v['detail']}{ap} | {v['pair']}")
                    for v in internalV:
                        icon = 'ℹ️ ' if v.get('note') else '⚠️ '
                        print(f"    {icon} [{v['fl']}] {v['detail']} | {v['pair']}")
                else:
                    print('✅ 이상없음')

                await page.wait_for_timeout(1200)

            except PWTimeout:
                print('⏱️ 타임아웃 (스킵)')
            except Exception as e:
                print(f'💥 오류: {e}')

        await browser.close()

    if not all_results:
        print(f'\n✅ {year}년 {month:02d}월 전체 위반사항 없음')
        input('\n엔터 누르면 종료...')
        return

    out_path, total_v, total_i = save_excel(all_results, year, month)

    print(f'\n{"="*50}')
    print(f'✅ 조회 완료: {len(all_results)}일 위반 발생 ({start_day}일~{last_day}일)')
    print(f'  🚨 규정위반: {total_v}건')
    print(f'  ⚠️  내부위반: {total_i}건')
    print(f'\n저장 완료: {out_path}')
    input('\n엔터 누르면 종료...')


if __name__ == '__main__':
    asyncio.run(main())
