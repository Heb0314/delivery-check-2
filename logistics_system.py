# Source Generated with Decompyle++
# File: logistics_system.pyc (Python 3.13)

'''
거리ㆍ단가 검증 시스템
실제 기능 구현 버전
'''
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import re
import os
from datetime import datetime
from playwright.sync_api import sync_playwright
import time

# 내장된 단가표 엑셀 파일 (exe와 같은 폴더의 embedded_pricing.xlsx)
# 원본 exe 실행 시에는 아래 바이트가 코드에 내장되어 있었음
def _load_embedded_xlsx():
    try:
        import sys
        if getattr(sys, 'frozen', False):
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        with open(os.path.join(base, 'embedded_pricing.xlsx'), 'rb') as _f:
            return _f.read()
    except FileNotFoundError:
        return b''
EMBEDDED_XLSX_BYTES = _load_embedded_xlsx()

class LogisticsSystem:
    def __init__(self, root):
        self.root = root
        self.root.title('거리ㆍ단가 검증 시스템')
        
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
            self.root.geometry('1000x750')
            self.root.minsize(900, 650)
            self.root.configure(bg = '#f5f5f7')
            self.input_file_path = None
            self.pricing_file_path = 'C:\\Users\\User\\Desktop\\실무역량\\퀵 이용단가.xlsx'
            self.input_df = None
            self.result_df = None
            self.pricing_area_df = None
            self.pricing_distance_df = None
            self.current_screen = 1
            self.main_container = tk.Frame(root, bg = '#f5f5f7')
            self.main_container.pack(fill = tk.BOTH, expand = True, padx = 20, pady = 20)
            self.show_step1_screen()
            return None
        except:
            pass


    
    def clear_main_container(self):
        '''메인 컨테이너 클리어'''
        for widget in self.main_container.winfo_children():
            widget.destroy()

    
    def preprocess_address(self, address):
        '''주소 전처리: 우편번호, 건물명 등 제거'''
        if bool(pd.isna(address)) or address == '':
            return ''
        address = str(address)
        address = re.sub('\\(\\d{5}\\)\\s*', '', address)
        address = re.sub('\\([^)]*\\)', '', address)
        building_keywords = [
            '주식회사',
            '회사',
            '빌딩',
            '타워',
            '센터',
            '아파트',
            '오피스텔',
            '오피스',
            '상가',
            '건물',
            '맨션',
            '빌라',
            '연립',
            '다세대',
            '단독',
            '공장',
            '사무실']
        for keyword in building_keywords:
            if not keyword in address:
                continue
            parts = address.split(keyword)
            address = parts[0]
            building_keywords
        address = re.sub('(\\d+(?:-\\d+)?)\\s*[가-힣]+.*$', '\\1', address)
        address = re.sub('\\s+', ' ', address)
        address = address.strip()
        return address

    
    def extract_district(self, address):
        '''주소에서 단가표 지역명 추출 (서울 구 / 경기 시·군 / 인천)'''
        if bool(pd.isna(address)) or address == '':
            return None
        address = re.sub('\\s+', ' ', str(address)).strip()
        names = []
        if self.pricing_area_df is not None and len(self.pricing_area_df) > 0:
            names = list(self.pricing_area_df['지역'])
        if '여의도' in address and '여의도동' in names:
            return '여의도동'
        if '인천' in address and '공항' in address and '인천국제공항' in names:
            return '인천국제공항'
        m = re.search('서울(?:특별시|시)?\\s+(\\S+구)', address)
        if bool(m):
            return m.group(1)
        m = re.search('경기(?:도)?\\s+(\\S+?[시군])(?=\\s|$)', address)
        if bool(m):
            return m.group(1)
        if bool(re.match('인천(?:광역시|시)?(?:\\s|$)', address)):
            return '인천시'
        if bool(re.match('(부산|대구|광주광역시|광주 |대전|울산|세종|강원|충청|충북|충남|전라|전북|전남|경상|경북|경남|제주)', address)):
            return None
        for name in names:
            if not bool(re.search('(^|\\s)' + re.escape(name) + '(\\s|$)', address)):
                continue
            names
            return name

    
    def load_pricing_table(self):
        """단가표 로드 - exe에 내장된 단가표 사용 (exe 옆에 '퀵 이용단가.xlsx'가 있으면 그 파일 우선)"""
        import io
        import sys

        try:
            if self.pricing_area_df is not None and self.pricing_distance_df is not None:
                return True

            def to_price(v):
                if pd.isna(v):
                    return None
                try:
                    return int(float(str(v).replace(',', '').replace('원', '').strip()))
                except Exception:
                    return None

            def parse(source):
                if isinstance(source, io.BytesIO):
                    source.seek(0)
                df1 = pd.read_excel(source, sheet_name=0, header=None)
                area_data = []
                for name_col in (1, 8):
                    for i in range(5, len(df1)):
                        name = df1.iat[i, name_col]
                        if pd.isna(name):
                            continue
                        name = str(name).strip().lstrip('※').strip()
                        if not name or name == '연계건':
                            continue
                        prices = [to_price(df1.iat[i, name_col + k]) for k in range(1, 6)]
                        if all(p is None for p in prices):
                            continue
                        area_data.append({
                            '지역': name,
                            '구분': '서울' if name_col == 1 else '경기',
                            '오토바이_편도': prices[0],
                            '오토바이_왕복': prices[1],
                            '다마스_편도': prices[2],
                            '다마스_왕복': prices[3],
                            '특수차량_편도': prices[4]})
                area_df = pd.DataFrame(area_data)
                if isinstance(source, io.BytesIO):
                    source.seek(0)
                df2 = pd.read_excel(source, sheet_name=1, header=None)
                dist_data = []
                for i in range(len(df2)):
                    label = df2.iat[i, 0]
                    if pd.isna(label) or 'km' not in str(label):
                        continue
                    label = str(label).strip()
                    m = re.search(r'(\d+)km이상~(\d+)km미만', label)
                    if m:
                        lo, hi = int(m.group(1)), int(m.group(2))
                        dist_data.append({
                            '거리': label,
                            'min_km': lo,
                            'max_km': hi,
                            '오토바이': to_price(df2.iat[i, 1]),
                            '다마스': to_price(df2.iat[i, 2]),
                            '특수차량': to_price(df2.iat[i, 3])})
                if len(area_data) == 0 or len(dist_data) == 0:
                    raise ValueError('단가표 양식을 인식하지 못했습니다')
                return area_df, pd.DataFrame(dist_data)

            tables = None
            try:
                exe_dir = os.path.dirname(os.path.abspath(sys.executable))
                external = os.path.join(exe_dir, '퀵 이용단가.xlsx')
                if os.path.exists(external):
                    tables = parse(external)
            except Exception:
                tables = None

            if tables is None:
                tables = parse(io.BytesIO(EMBEDDED_XLSX_BYTES))
            self.pricing_area_df, self.pricing_distance_df = tables
            return True

        except Exception as e:
            self.pricing_area_df = None
            self.pricing_distance_df = None
            messagebox.showerror('오류', '단가표 로드 실패:\n' + str(e))
            return False


    def get_price_from_area(self, district, vehicle_type, trip_type):
        '''Sheet 1: 지역별 단가 조회'''
        if self.pricing_area_df is None or district is None:
            return None
        try:
            row = self.pricing_area_df[self.pricing_area_df['지역'] == district]
            if row.empty:
                return None
            try:
                col_suffix = '_편도' if trip_type == '편도' else '_왕복'
                col_name = vehicle_type + col_suffix
                if col_name in row.columns:
                    price = row[col_name].values[0]
                    if pd.notna(price):
                        return int(price)
                return None
            except Exception:
                return None
        except Exception:
            return None



    
    def get_price_from_distance(self, distance, vehicle_type, trip_type):
        '''Sheet 2: 거리별 단가 조회 (소수점 0.0~0.4 버림, 0.5~0.9 올림)'''
        import math
        if self.pricing_distance_df is None or distance is None:
            return None
        try:
            d = math.floor(float(distance) + 0.5)
            for _, row in self.pricing_distance_df.iterrows():
                if row['min_km'] <= d < row['max_km']:
                    if vehicle_type not in row.index:
                        return None
                    price = row[vehicle_type]
                    if price is None or pd.isna(price):
                        return None
                    price = int(price)
                    return price * (2 if trip_type == '왕복' else 1)
            return None
        except Exception:
            return None


    def calculate_distance_naver(self, origin, destination, diag = None):
        '''네이버 지도(자동차) 최단거리 조회 - Edge 헤드리스 (스레드별 브라우저).
origin=None 으로 호출하면 현재 스레드의 브라우저를 종료한다.
성공 시 km(float), 실패 시 None. 실패 사유는 반환 직전 self._naver_errors[스레드]'''
        import sys
        import json
        import threading
        import re
        import os
        import time
        tid = threading.get_ident()
        if not bool(hasattr(self, '_naver_states')):
            self._naver_states = { }
            self._naver_errors = { }
            self._naver_lock = threading.Lock()
        states = self._naver_states
        # origin=None: 현재 스레드의 브라우저 종료
        if origin is None:
            st = states.pop(tid, None)
            if st:
                try:
                    if st.get('page'):
                        st['page'].close()
                except Exception:
                    pass
                try:
                    if st.get('context'):
                        st['context'].close()
                except Exception:
                    pass
                try:
                    if st.get('browser'):
                        st['browser'].close()
                except Exception:
                    pass
                try:
                    if st.get('pw'):
                        st['pw'].stop()
                except Exception:
                    pass
            return None
        NAVER_URL = os.environ.get('DISTANCE_NAVER_URL', 'https://map.naver.com/p/directions/-/-/-/car?c=15,0,0,0,dh')
        exe_dir = os.path.dirname(os.path.abspath(sys.executable)) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        log_dir = os.path.join(exe_dir, '거리조회_로그')
        try:
            os.makedirs(log_dir, exist_ok=True)
        except Exception:
            pass
        def _log(msg):
            try:
                with open(os.path.join(log_dir, 'debug.txt'), 'a', encoding='utf-8') as f:
                    f.write(time.strftime('%Y-%m-%d %H:%M:%S') + ' ' + msg + '\n')
            except Exception:
                pass
        def _fail(reason):
            self._naver_errors[tid] = reason
            _log('FAIL %s -> %s : %s' % (origin, destination, reason))
            return None
        # 스레드별 브라우저 준비
        st = states.get(tid)
        if not st or not st.get('page'):
            try:
                from playwright.sync_api import sync_playwright
                pw = sync_playwright().start()
            except Exception as e:
                return _fail('브라우저 제어 모듈 시작 실패: ' + str(e))
            browser = None
            last_err = ''
            for ch in os.environ.get('DISTANCE_BROWSER', 'msedge,chrome').split(','):
                ch = ch.strip()
                if not ch:
                    continue
                try:
                    browser = pw.chromium.launch(channel=ch, headless=True, args=['--disable-blink-features=AutomationControlled'])
                    break
                except Exception as e:
                    last_err = str(e)
                    continue
            if browser is None:
                try:
                    pw.stop()
                except Exception:
                    pass
                return _fail('Edge 브라우저를 실행할 수 없습니다.\n' + last_err)
            try:
                context = browser.new_context(user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0', locale='ko-KR')
                # 이미지/폰트/광고 차단 (속도 향상)
                for pat in ('**/*.{png,jpg,jpeg,gif,webp,svg,woff,woff2,ttf}', '**/veta.naver.com/**', '**/ads/**'):
                    context.route(pat, lambda route: route.abort())
                page = context.new_page()
                st = {'pw': pw, 'browser': browser, 'context': context, 'page': page}
                states[tid] = st
            except Exception as e:
                try:
                    browser.close()
                except Exception:
                    pass
                try:
                    pw.stop()
                except Exception:
                    pass
                return _fail('브라우저 초기화 실패: ' + str(e))
        page = st['page']
        try:
            page.goto(NAVER_URL, timeout=30000, wait_until='domcontentloaded')
        except Exception as e:
            return _fail('네이버 지도 접속 실패: ' + str(e))
        # 출발지 입력
        try:
            start_input = page.locator('.search_input_box_wrap.start input, input[id^="input_search"]').first
            start_input.click(timeout=10000)
            start_input.fill('')
            start_input.type(origin, delay=30)
            page.wait_for_timeout(1200)
            # 자동완성 첫 번째 항목 선택
            opt = page.locator('[role="listbox"]:not(.list_search_result) [role="option"], ul.list_search_result > li[role="option"], ul.list_search_result > li').first
            try:
                opt.click(timeout=5000)
            except Exception:
                start_input.press('Enter')
            page.wait_for_timeout(800)
        except Exception as e:
            return _fail('출발지 입력 실패: ' + str(e))
        # 도착지 입력
        try:
            goal_input = page.locator('.search_input_box_wrap.goal input').first
            goal_input.click(timeout=10000)
            goal_input.fill('')
            goal_input.type(destination, delay=30)
            page.wait_for_timeout(1200)
            opt = page.locator('[role="listbox"]:not(.list_search_result) [role="option"], ul.list_search_result > li[role="option"], ul.list_search_result > li').first
            try:
                opt.click(timeout=5000)
            except Exception:
                goal_input.press('Enter')
            page.wait_for_timeout(1500)
        except Exception as e:
            return _fail('도착지 입력 실패: ' + str(e))
        # 경로 목록에서 거리우선 km 추출
        try:
            # 경로 결과가 나타날 때까지 대기
            page.wait_for_timeout(2500)
            body_text = page.locator('body').inner_text(timeout=10000)
        except Exception as e:
            return _fail('경로 결과 읽기 실패: ' + str(e))
        # 경로 파싱: 번호 / 경로명(거리우선 등) / 소요시간 / 거리
        route_re = re.compile(r'(?:^|\n)\s*(\d{1,2})\s*([^\n]{1,40})\n\s*((?:\d+\s*시간\s*)?(?:\d+\s*분)?)\s*(\d{1,4}(?:[.,]\d+)?)\s*(km|m)\s*(?:\n|$)')
        best_km = None
        for m in route_re.finditer(body_text):
            idx, name, _tm, dist, unit = m.group(1), m.group(2).strip(), m.group(3), m.group(4), m.group(5)
            try:
                km = float(dist.replace(',', ''))
            except Exception:
                continue
            if unit == 'm':
                km = km / 1000.0
            # 거리우선 경로를 우선 선택
            if '거리우선' in name:
                best_km = km
                break
            if best_km is None:
                best_km = km
        if best_km is None:
            # 대체 파싱: "123.4km" 형태 직접 탐색
            m2 = re.search(r'(\d{1,4}(?:\.\d+)?)\s*km', body_text)
            if m2:
                try:
                    best_km = float(m2.group(1))
                except Exception:
                    pass
        if best_km is None:
            return _fail('경로 km 파싱 실패')
        _log('OK %s -> %s : %.1f km' % (origin, destination, best_km))
        return float(best_km)

    
    def select_file(self):
        '''파일 선택 대화상자'''
        file_path = filedialog.askopenfilename(title = '엑셀 파일 선택', filetypes = [
            ('Excel files', '*.xlsx *.xls')])
        if bool(file_path):
            self.input_file_path = file_path
            self.update_file_display(os.path.basename(file_path))
            return None

    
    def update_file_display(self, filename):
        '''선택된 파일 표시 업데이트'''
        for widget in self.file_display_area.winfo_children():
            widget.destroy()
        file_card = tk.Frame(self.file_display_area, bg = '#e8f4fd', relief = tk.SOLID, bd = 1)
        file_card.pack(fill = tk.X, pady = 5)
        left_frame = tk.Frame(file_card, bg = '#e8f4fd')
        left_frame.pack(side = tk.LEFT, fill = tk.BOTH, expand = True, padx = 10, pady = 8)
        tk.Label(left_frame, text = '📄', font = ('맑은 고딕', 14), bg = '#e8f4fd').pack(side = tk.LEFT, padx = (0, 8))
        tk.Label(left_frame, text = filename, font = ('맑은 고딕', 9, 'bold'), bg = '#e8f4fd', fg = '#0071e3').pack(side = tk.LEFT)
        remove_btn = tk.Button(file_card, text = '✕', font = ('맑은 고딕', 10, 'bold'), bg = '#e8f4fd', fg = '#666', bd = 0, cursor = 'hand2', command = self.remove_file)
        remove_btn.pack(side = tk.RIGHT, padx = 8)
        if not bool(hasattr(self, 'file_hint_shown')):
            hint_label = tk.Label(self.file_display_area, text = "※ 다른 파일을 선택하려면 '파일 선택' 버튼을 다시 클릭하세요", font = ('맑은 고딕', 8), bg = '#fafafa', fg = '#999')
            hint_label.pack(pady = (5, 0))
            self.file_hint_shown = True
            return None

    
    def remove_file(self):
        '''파일 선택 취소'''
        self.input_file_path = None
        for widget in self.file_display_area.winfo_children():
            widget.destroy()
        tk.Label(self.file_display_area, text = '선택된 파일 없음', font = ('맑은 고딕', 9), bg = '#fafafa', fg = '#999').pack(pady = 10)

    
    def process_data(self):
        '''데이터 처리 및 계산 실행 (단가표는 퀵서비스일 때만 적용, 거리는 네이버 병렬 조회)'''
        import threading
        if not bool(self.input_file_path):
            messagebox.showerror('오류', '엑셀 파일을 먼저 선택해주세요.')
            return None
        
        def read_via_excel(path):
            import subprocess
            import base64
            ps = '\n$ErrorActionPreference = \'Stop\'\n$p = $env:DIST_XL_PATH\n$xl = $null; $wb = $null; $failed = $false\ntry {\n  $xl = New-Object -ComObject Excel.Application\n  $xl.Visible = $false\n  $xl.DisplayAlerts = $false\n  $xl.AskToUpdateLinks = $false\n  $wb = $xl.Workbooks.Open($p, 0, $true)\n  $ws = $wb.Worksheets.Item(1)\n  $ur = $ws.UsedRange\n  $nr = $ur.Rows.Count\n  $nc = $ur.Columns.Count\n  $v = $ur.Value2\n  $inv = [Globalization.CultureInfo]::InvariantCulture\n  $sb = New-Object System.Text.StringBuilder\n  [void]$sb.AppendLine("SIZE`t$nr`t$nc")\n  for ($r = 1; $r -le $nr; $r++) {\n    for ($c = 1; $c -le $nc; $c++) {\n      if ($nr -eq 1 -and $nc -eq 1) { $x = $v } else { $x = $v[$r, $c] }\n      if ($x -ne $null) {\n        if ($x -is [double]) { $t = \'N\'; $s = $x.ToString(\'R\', $inv) }\n        else { $t = \'S\'; $s = [string]$x }\n        $b = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($s))\n        [void]$sb.AppendLine("$r`t$c`t$t`t$b")\n      }\n    }\n  }\n  [Console]::Out.Write($sb.ToString())\n} catch {\n  $failed = $true\n  [Console]::Out.Write("ERROR`t" + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($_.Exception.Message)))\n} finally {\n  if ($wb -ne $null) { try { $wb.Close($false) } catch {} }\n  if ($xl -ne $null) { try { $xl.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }\n}\nif ($failed) { exit 1 }\n'
            enc = base64.b64encode(ps.encode('utf-16-le')).decode('ascii')
            env = dict(os.environ)
            env['DIST_XL_PATH'] = os.path.abspath(path)
            r = subprocess.run([
                'powershell.exe',
                '-NoProfile',
                '-NonInteractive',
                '-ExecutionPolicy',
                'Bypass',
                '-EncodedCommand',
                enc], stdin = subprocess.DEVNULL, stdout = subprocess.PIPE, stderr = subprocess.PIPE, env = env, timeout = 600, creationflags = 134217728)
            out = r.stdout.decode('ascii', 'ignore')
            if out.startswith('ERROR\t'):
                raise RuntimeError(base64.b64decode(out.split('\t', 1)[1].strip()).decode('utf-8', 'replace'))
            lines = [ln for ln in out.splitlines() if ln.strip()]
            if not lines or not lines[0].startswith('SIZE\t'):
                err = r.stderr.decode('cp949', 'replace').strip()
                raise RuntimeError(err[:300] or '엑셀 실행 결과가 없습니다 (엑셀 미설치?)')
            _, nr, nc = lines[0].split('\t')
            nr = int(nr)
            nc = int(nc)
            grid = [[None] * nc for _ in range(nr)]
            for ln in lines[1:]:
                parts = ln.split('\t')
                if len(parts) != 4:
                    continue
                rr, cc, t, b = int(parts[0]), int(parts[1]), parts[2], parts[3]
                s = base64.b64decode(b).decode('utf-8', 'replace')
                if t == 'N':
                    fv = float(s)
                    val = int(fv) if fv.is_integer() else fv
                else:
                    val = s
                grid[rr - 1][cc - 1] = val
            if nr == 0:
                return pd.DataFrame()
            headers = []
            for i, h in enumerate(grid[0]):
                name = str(h).strip() if h is not None and str(h).strip() else 'Unnamed: %d' % i
                while name in headers:
                    name = name + '_'
                headers.append(name)
            rows = [rw for rw in grid[1:] if any(x is not None and str(x).strip() != "" for x in rw)]
            df = pd.DataFrame(rows, columns=headers)
            for col in df.columns:
                try:
                    df[col] = pd.to_numeric(df[col])
                except Exception:
                    pass
            return df

        CAMPUS_QUERY = '서울특별시 서초구 성촌길 56'
        
        def is_campus(addr):
            if addr is None or pd.isna(addr):
                return False
            t = re.sub('\\s+', '', str(addr)).lower()
            keys = ('r&d캠퍼스', 'rnd캠퍼스', '알앤디캠퍼스', 'r&dcampus', 'rndcampus', '성촌길33', '성촌길34', '성촌길56')
            if any(k in t for k in keys):
                return True
            return re.search('우면동(678|679|689)(?!\\d)', t) is not None

        
        def naver_query(raw, fallback):
            """네이버 검색용 주소: '도로명 + 건물번호' 또는 '동/리 + 지번'까지만 남김"""
            if raw is None or pd.isna(raw):
                return fallback
            a = re.sub('\\(\\d{5}\\)', ' ', str(raw))
            a = re.sub('\\([^)]*\\)', ' ', a)
            a = re.sub('\\s+', ' ', a).strip()
            toks = a.split(' ')
            for i in range(len(toks) - 1):
                if not bool(re.fullmatch('[가-힣][가-힣0-9]*(로|길)', toks[i])):
                    continue
                m = re.match('(\\d+(?:-\\d+)?)', toks[i + 1])
                if not m:
                    continue
                return ' '.join(toks[:i + 1]) + ' ' + m.group(1)
            for i in range(len(toks) - 1):
                if not re.fullmatch('[가-힣][가-힣0-9]*(동|리|가)', toks[i]):
                    continue
                m = re.match('(\\d+(?:-\\d+)?)', toks[i + 1])
                if not m:
                    continue
                return ' '.join(toks[:i + 1]) + ' ' + m.group(1)
            return fallback or a

        
        def classify(row):
            '''차종·편도/왕복은 전달내용에서 자동 판별 (GUI 선택값 미사용)'''
            reason = ''
            if '전달내용' in self.input_df.columns and pd.notna(row['전달내용']):
                reason = str(row['전달내용'])
            trip = '왕복' if '왕복' in reason else '편도'
            if '연계' in reason:
                vehicle = '연계'
            elif re.search(r'1\s*[tT]', reason):
                vehicle = '1t'
            elif '차량' in reason or '직송' in reason:
                vehicle = '다마스'
            else:
                vehicle = '오토바이'
            return (vehicle, trip, '일반', None)



        try:
            category = self.category_combo.get()
        except Exception:
            category = '퀵서비스'
        use_pricing = category == '퀵서비스'
        if use_pricing and not self.load_pricing_table():
            return None
        progress_window = None
        try:
            self.input_df = pd.read_excel(self.input_file_path)
        except Exception as first_error:
            try:
                self.input_df = read_via_excel(self.input_file_path)
            except Exception as second_error:
                raise RuntimeError('엑셀 파일을 읽을 수 없습니다.\n- 기본 읽기: ' + str(first_error) + '\n- 엑셀 프로그램으로 읽기: ' + str(second_error) + '\n\n문서보안이 걸린 파일이면 보안 해제 후 다시 시도해 주세요.')
        self.input_df = self.input_df.drop(columns=[c for c in ('계산거리(km)', '단가표금액', '차액', '검증결과', '적용기준', '업체기재거리(km)') if c in self.input_df.columns])
        required_columns = ['전달내용', '출발지', '도착지', '비용']
        missing_columns = [col for col in required_columns if col not in self.input_df.columns]
        if missing_columns:
            messagebox.showerror('오류', '필수 컬럼이 없습니다: ' + ', '.join(missing_columns))
            return None
        progress_window = tk.Toplevel(self.root)
        progress_window.title('처리 중...')
        progress_window.geometry('450x180')
        progress_window.configure(bg='#f5f5f7')
        progress_label = tk.Label(progress_window, text='데이터 처리 중...', font=('맑은 고딕', 11, 'bold'), bg='#f5f5f7')
        progress_label.pack(pady=20)
        progress_bar = ttk.Progressbar(progress_window, length=350, mode='determinate')
        progress_bar.pack(pady=12)
        progress_detail = tk.Label(progress_window, text='0% (0/0)', font=('맑은 고딕', 10), bg='#f5f5f7', fg='#6e6e73')
        progress_detail.pack(pady=8)
        progress_window.transient(self.root)
        progress_window.grab_set()
        progress_window.protocol('WM_DELETE_WINDOW', lambda: None)
        progress_window.update()
        total_rows = len(self.input_df)
        plans = []
        needed = []
        needed_set = set()
        for idx, row in self.input_df.iterrows():
            origin = self.preprocess_address(row['출발지'])
            destination = self.preprocess_address(row['도착지'])
            vehicle_type, trip_type, mode, claimed = classify(row)
            p = {'vehicle': vehicle_type, 'trip': trip_type, 'mode': mode, 'claimed': claimed, 'cost': row['비용'], 'price': None, 'basis': None, 'key': None, 'status': None}
            if not origin and not destination:
                p['status'] = '제외(합계 등)'
                plans.append(p)
                continue
            if not use_pricing:
                p['status'] = '미적용(용차)'
                plans.append(p)
                continue
            if vehicle_type == '연계':
                p['status'] = '연계(검증제외)'
                p['price'] = p['cost'] if pd.notna(p['cost']) else None
                p['basis'] = '연계 건(%s)' % trip_type
                plans.append(p)
                continue
            price_vehicle = '특수차량' if vehicle_type == '1t' else vehicle_type
            p['price_vehicle'] = price_vehicle
            o_campus = is_campus(row['출발지'])
            d_campus = is_campus(row['도착지'])
            area_district, area_side = None, ''
            if o_campus and not d_campus:
                area_district = self.extract_district(destination)
                area_side = '캠퍼스 출발'
            elif d_campus and not o_campus:
                area_district = self.extract_district(origin)
                area_side = '캠퍼스 도착'
            price = None
            if mode != '직송':
                price = self.get_price_from_area(area_district, p.get('price_vehicle', vehicle_type), trip_type)
            if price is None:
                if not origin or not destination:
                    p['status'] = '주소 없음'
                else:
                    q_o = CAMPUS_QUERY if o_campus else naver_query(row['출발지'], origin)
                    q_d = CAMPUS_QUERY if d_campus else naver_query(row['도착지'], destination)
                    key = (q_o, q_d)
                    if re.sub('\\s+', '', q_o) == re.sub('\\s+', '', q_d):
                        p['status'] = '불일치'
                        p['basis'] = '수동 확인 필요(출발지와 도착지 주소가 같음)'
                        plans.append(p)
                        continue
                    p['key'] = key
                    if key not in needed_set:
                        needed_set.add(key)
                        needed.append(key)
            else:
                if mode == '긴급':
                    price *= 2
                p['price'] = price
                try:
                    side = self.pricing_area_df.loc[self.pricing_area_df['지역'] == area_district, '구분'].iloc[0]
                except Exception:
                    side = ''
                p['dist_text'] = (side + ' 단가적용') if side else '지역 단가적용'
                p['basis'] = f'지역단가({area_district}, {area_side}, {vehicle_type}-{trip_type})'
            plans.append(p)
        DIAG_ROUTES = [
            ('서울특별시 서초구 성촌길 56', '경기도 수원시 영통구 삼성로 129'),
            ('서울특별시 서초구 성촌길 56', '대전광역시 유성구 대학로 291'),
            ('서울특별시 중구 장충단로 60', '서울특별시 중구 동호로 249'),
            ('경기도 수원시 영통구 삼성로 129', '경기도 성남시 중원구 상대원동 5469'),
            ('서울특별시 서초구 성촌길 56', '부산광역시 강서구 녹산산단262로13번길 12'),
        ]
        import sys as _sys0
        _diag_flag = os.path.exists(os.path.join(os.path.dirname(os.path.abspath(_sys0.executable)), '진단.txt'))
        if use_pricing and (_diag_flag or os.environ.get('DISTANCE_DIAG') == '1'):
            diag_lines = []
            for i_, (d_o, d_d) in enumerate(DIAG_ROUTES):
                progress_label.config(text='네이버 지도 진단 중... (%d/%d)' % (i_ + 1, len(DIAG_ROUTES)))
                progress_detail.config(text='실제 화면을 단계별로 기록하고 있습니다')
                progress_window.update()
                km_ = self.calculate_distance_naver(d_o, d_d, diag=True)
                err_ = self._naver_errors.get(threading.get_ident(), '')
                diag_lines.append(f'{d_o} -> {d_d} : {("%.1f km" % km_) if km_ is not None else "실패"} {err_}')
                progress_window.update()
            self.calculate_distance_naver(None)
            import sys as _sys
            ld = os.path.join(os.path.dirname(os.path.abspath(_sys.executable)), '거리조회_로그')
            os.makedirs(ld, exist_ok=True)
            with open(os.path.join(ld, '진단_요약.txt'), 'w', encoding='utf-8') as f:
                f.write('\n'.join(diag_lines) + '\n')
        results = {}
        if needed:
            n_workers = int(os.environ.get('DISTANCE_WORKERS', '3'))
            n_workers = max(1, min(n_workers, len(needed)))
            queue = list(needed)
            qlock = threading.Lock()
            done = [0]

            def worker():
                while True:
                    with qlock:
                        if not queue:
                            try:
                                self.calculate_distance_naver(None)
                            except Exception:
                                pass
                            return
                        key = queue.pop(0)
                    km = None
                    err = ''
                    try:
                        km = self.calculate_distance_naver(key[0], key[1])
                        if km is None:
                            err = self._naver_errors.get(threading.get_ident(), '')
                    except Exception as ex:
                        err = str(ex)
                    with qlock:
                        results[key] = (km, err)
                        done[0] += 1

            threads = [threading.Thread(target=worker, daemon=True) for _ in range(n_workers)]
            for t in threads:
                t.start()
            while any(t.is_alive() for t in threads):
                d = done[0]
                pct = int(d / len(needed) * 100)
                progress_bar['value'] = pct
                progress_label.config(text='네이버 지도 거리 조회 중... (동시 %d개)' % n_workers)
                progress_detail.config(text=f'{pct}% ({d}/{len(needed)}건, 중복 구간 제외)')
                progress_window.update()
                time.sleep(0.15)
        progress_label.config(text='금액 계산 중...')
        progress_window.update()
        distances = []
        calculated_prices = []
        differences = []
        validations = []
        bases = []
        fail_count = 0
        fatal_msg = None
        for p in plans:
            distance = None
            price = p['price']
            status = p['status']
            basis = p['basis']
            if p['key'] is not None:
                km, err = results.get(p['key'], (None, '조회 안 됨'))
                if km is None:
                    fail_count += 1
                    status = '거리 조회 실패'
                    if err and ('브라우저' in err or 'Edge' in err):
                        fatal_msg = err
                else:
                    distance = km
                    price = self.get_price_from_distance(km, p.get('price_vehicle', p['vehicle']), p['trip'])
                    basis = f"거리단가({p['vehicle']}-{p['trip']})"
                    if price is None:
                        status = '단가 없음'
                    elif p['mode'] == '긴급':
                        price *= 2
            distances.append(distance if distance is not None else p.get('dist_text'))
            if p.get('vehicle') == '연계':
                calculated_prices.append(price)
                differences.append(0)
                validations.append('연계(검증제외)')
                bases.append(basis)
                continue
            if price is None or status is None:
                calculated_prices.append(None)
                differences.append(None)
                validations.append(status or '단가 없음')
                bases.append(basis)
                continue
            calculated_prices.append(price)
            original_price = p['cost'] if pd.notna(p['cost']) else 0
            try:
                diff = price - float(original_price)
            except Exception:
                diff = None
            differences.append(diff)
            validations.append('정상' if diff == 0 else '불일치')
            bases.append(basis)
        progress_bar['value'] = 100
        progress_detail.config(text='100% (%d/%d)' % (total_rows, total_rows))
        progress_window.update()
        cost_col_idx = self.input_df.columns.get_loc('비용')
        self.result_df = self.input_df.copy()
        self.result_df.insert(cost_col_idx + 1, '계산거리(km)', pd.array([d if isinstance(d, str) else (None if d is None else float(d)) for d in distances], dtype=object))
        self.result_df.insert(cost_col_idx + 2, '단가표금액', pd.array(calculated_prices, dtype='Int64'))
        self.result_df.insert(cost_col_idx + 3, '차액', pd.array([None if d is None else int(round(d)) for d in differences], dtype='Int64'))
        self.result_df.insert(cost_col_idx + 4, '검증결과', validations)
        self.result_df.insert(cost_col_idx + 5, '적용기준', bases)
        self.result_df.insert(cost_col_idx + 6, '차종', [p.get('vehicle', '') for p in plans])
        self.result_df.insert(cost_col_idx + 7, '편도/왕복', [p.get('trip', '') for p in plans])
        progress_window.destroy()
        progress_window = None
        if fail_count:
            msg = f'{fail_count}건의 네이버 지도 거리 조회에 실패했습니다.\n'
            if fatal_msg:
                msg += '\n' + fatal_msg + '\n'
            msg += "\n실행 파일 옆 '거리조회_로그' 폴더에 원인 기록이 저장되었습니다."
            messagebox.showwarning('거리 조회', msg)
        self.show_step2_screen()
        return None

    
    def save_result(self):
        '''결과 저장: [검증결과] 시트(화면과 동일, 불일치 빨간색) + [전체데이터] 시트(원본 열 포함)'''
        if self.result_df is None:
            messagebox.showerror('오류', '저장할 결과가 없습니다.')
            return None

        try:
            base_name = os.path.splitext(os.path.basename(self.input_file_path))[0]
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_filename = base_name + '_결과_' + timestamp + '.xlsx'
            output_path = filedialog.asksaveasfilename(
                title='결과 파일 저장',
                initialfile=output_filename,
                defaultextension='.xlsx',
                filetypes=[('Excel files', '*.xlsx')])
            if not output_path:
                return None

            # Excel 저장 (openpyxl 스타일 적용)
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                self.result_df.to_excel(writer, sheet_name='검증결과', index=False)
                if self.input_df is not None:
                    self.input_df.to_excel(writer, sheet_name='전체데이터', index=False)

            # 불일치 셀 빨간색 표시
            try:
                from openpyxl import load_workbook
                from openpyxl.styles import PatternFill, Alignment
                wb = load_workbook(output_path)
                ws = wb['검증결과']
                red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
                # 검증결과 열에서 불일치 행 표시 (간략화)
                wb.save(output_path)
            except Exception:
                pass

            messagebox.showinfo('성공', f'결과 파일이 저장되었습니다:\n{output_path}')
        except PermissionError:
            messagebox.showerror('오류', '파일 저장 중 오류 발생:\n같은 이름의 파일이 엑셀에서 열려 있으면 닫고 다시 저장해 주세요.')
        except Exception as e:
            messagebox.showerror('오류', '파일 저장 중 오류 발생:\n' + str(e))


    def on_category_change(self, event = None):
        '''배송 구분 변경 시 차종 옵션 업데이트'''
        category = self.category_combo.get()
        self.update_vehicle_options(category)

    
    def update_vehicle_options(self, category):
        '''선택한 구분에 따라 차종 옵션 표시'''
        for widget in self.vehicle_container.winfo_children():
            widget.destroy()
        tk.Label(self.vehicle_container, text = '차종:', font = ('맑은 고딕', 10), bg = 'white').grid(row = 0, column = 0, sticky = 'w', padx = (0, 10))
        if category == '퀵서비스':
            label_frame = tk.Frame(self.vehicle_container, bg = 'white')
            label_frame.grid(row = 0, column = 1, sticky = 'w')
            tk.Label(label_frame, text = '[ 퀵서비스 ]', font = ('맑은 고딕', 9, 'bold'), bg = 'white', fg = '#0071e3').pack(side = tk.LEFT, padx = (0, 10))
            self.vehicle_combo = ttk.Combobox(label_frame, values = [
                '오토바이',
                '다마스',
                '특수차량'], font = ('맑은 고딕', 10), state = 'readonly', width = 11)
            self.vehicle_combo.current(0)
            self.vehicle_combo.pack(side = tk.LEFT, padx = (0, 10))
            self.trip_type_combo = ttk.Combobox(label_frame, values = [
                '편도',
                '왕복'], font = ('맑은 고딕', 10), state = 'readonly', width = 11)
            self.trip_type_combo.current(0)
            self.trip_type_combo.pack(side = tk.LEFT)
            return None
        label_frame = tk.Frame(self.vehicle_container, bg = 'white')
        label_frame.grid(row = 0, column = 1, sticky = 'w')
        tk.Label(label_frame, text = '[ 용차 ]', font = ('맑은 고딕', 9, 'bold'), bg = 'white', fg = '#0066cc').pack(side = tk.LEFT, padx = (0, 10))
        self.tonnage_combo = ttk.Combobox(label_frame, values = [
            '1톤',
            '5톤'], font = ('맑은 고딕', 10), state = 'readonly', width = 11)
        self.tonnage_combo.current(0)
        self.tonnage_combo.pack(side = tk.LEFT, padx = (0, 10))
        self.vehicle_type_combo = ttk.Combobox(label_frame, values = [
            '탑차',
            '윙바디'], font = ('맑은 고딕', 10), state = 'readonly', width = 11)
        self.vehicle_type_combo.current(0)
        self.vehicle_type_combo.pack(side = tk.LEFT)

    
    def show_step1_screen(self):
        '''1단계: 데이터 입력 화면'''
        self.clear_main_container()
        self.current_screen = 1
        title_frame = tk.Frame(self.main_container, bg = '#f5f5f7')
        title_frame.pack(fill = tk.X, pady = (0, 20))
        tk.Label(title_frame, text = '거리ㆍ단가 검증 시스템', font = ('맑은 고딕', 18, 'bold'), bg = '#f5f5f7', fg = '#1d1d1f').pack()
        tk.Label(title_frame, text = '1단계: 데이터 입력 및 설정', font = ('맑은 고딕', 11), bg = '#f5f5f7', fg = '#6e6e73').pack()
        input_frame = tk.Frame(self.main_container, bg = 'white', relief = tk.RAISED, bd = 2)
        input_frame.pack(fill = tk.BOTH, expand = True, padx = 10, pady = 10)
        section1 = tk.LabelFrame(input_frame, text = '1. 배송 구분', font = ('맑은 고딕', 10, 'bold'), bg = 'white', fg = '#1d1d1f', padx = 15, pady = 10)
        section1.pack(fill = tk.X, padx = 15, pady = (15, 8))
        tk.Label(section1, text = '구분:', font = ('맑은 고딕', 10), bg = 'white').grid(row = 0, column = 0, sticky = 'w', padx = (0, 10))
        self.category_combo = ttk.Combobox(section1, values = [
            '퀵서비스',
            '용차'], font = ('맑은 고딕', 10), state = 'readonly', width = 25)
        self.category_combo.current(0)
        self.category_combo.grid(row = 0, column = 1, sticky = 'w')
        self.category_combo.bind('<<ComboboxSelected>>', self.on_category_change)
        self.section2 = tk.LabelFrame(input_frame, text = '2. 차종 세부선택', font = ('맑은 고딕', 10, 'bold'), bg = 'white', fg = '#1d1d1f', padx = 15, pady = 10)
        self.section2.pack(fill = tk.X, padx = 15, pady = 8)
        self.vehicle_container = tk.Frame(self.section2, bg = 'white')
        self.vehicle_container.pack(fill = tk.X, pady = (10, 5))
        self.update_vehicle_options('퀵서비스')
        section3 = tk.LabelFrame(input_frame, text = '3. 엑셀 파일 업로드', font = ('맑은 고딕', 10, 'bold'), bg = 'white', fg = '#1d1d1f', padx = 15, pady = 10)
        section3.pack(fill = tk.BOTH, expand = True, padx = 15, pady = 8)
        upload_area = tk.Frame(section3, bg = '#fafafa', relief = tk.GROOVE, bd = 2)
        upload_area.pack(fill = tk.BOTH, expand = True, pady = (8, 8))
        top_area = tk.Frame(upload_area, bg = '#fafafa')
        top_area.pack(fill = tk.X, pady = (15, 10))
        tk.Label(top_area, text = '📁', font = ('맑은 고딕', 28), bg = '#fafafa').pack()
        tk.Label(top_area, text = '출발지/도착지 데이터가 담긴 엑셀 파일', font = ('맑은 고딕', 9), bg = '#fafafa', fg = '#6e6e73').pack(pady = (5, 10))
        tk.Button(top_area, text = '파일 선택', font = ('맑은 고딕', 10, 'bold'), bg = '#0071e3', fg = 'white', padx = 25, pady = 7, cursor = 'hand2', relief = tk.FLAT, command = self.select_file).pack()
        self.file_display_area = tk.Frame(upload_area, bg = '#fafafa')
        self.file_display_area.pack(fill = tk.BOTH, expand = True, padx = 15, pady = (10, 15))
        tk.Label(self.file_display_area, text = '선택된 파일 없음', font = ('맑은 고딕', 9), bg = '#fafafa', fg = '#999').pack(pady = 10)
        button_frame = tk.Frame(self.main_container, bg = '#f5f5f7')
        button_frame.pack(fill = tk.X, pady = 15)
        tk.Button(button_frame, text = '데이터 처리 시작 ▶', font = ('맑은 고딕', 12, 'bold'), bg = '#0071e3', fg = 'white', padx = 40, pady = 12, cursor = 'hand2', relief = tk.RAISED, bd = 3, command = self.process_data).pack()

    
    def show_step2_screen(self):
        '''2단계: 결과 화면'''
        self.clear_main_container()
        self.current_screen = 2
        title_frame = tk.Frame(self.main_container, bg = '#f5f5f7')
        title_frame.pack(fill = tk.X, pady = (0, 20))
        tk.Label(title_frame, text = '계산 결과 및 검증', font = ('맑은 고딕', 18, 'bold'), bg = '#f5f5f7', fg = '#1d1d1f').pack()
        tk.Label(title_frame, text = '2단계: 거리 계산 및 금액 산출 결과', font = ('맑은 고딕', 11), bg = '#f5f5f7', fg = '#6e6e73').pack()
        result_frame = tk.Frame(self.main_container, bg = 'white', relief = tk.RAISED, bd = 2)
        result_frame.pack(fill = tk.BOTH, expand = True, padx = 10, pady = 10)
        if self.result_df is not None:
            total_count = len(self.result_df)
            dist_num = pd.to_numeric(self.result_df['계산거리(km)'], errors='coerce')
            amt_num = pd.to_numeric(self.result_df['단가표금액'], errors='coerce')
            n_dist = int(dist_num.notna().sum())
            total_distance = float(dist_num.sum()) if n_dist else 0
            total_amount = int(amt_num.sum()) if bool(amt_num.notna().any()) else 0
            avg_text = '%.1f km' % (total_distance / n_dist) if n_dist else '-'
            n_mis = int((self.result_df['검증결과'] == '불일치').sum()) if '검증결과' in self.result_df.columns else 0
            summary_frame = tk.Frame(result_frame, bg='#f5f5f7', relief=tk.SOLID, bd=1)
            summary_frame.pack(fill=tk.X, padx=20, pady=(20, 10))
            summary_grid = tk.Frame(summary_frame, bg='#f5f5f7')
            summary_grid.pack(pady=15, padx=20)
            summaries = [
                ('총 건수', f'{total_count}건', '#0071e3'),
                ('총 거리 (거리조회 %d건)' % n_dist, f'{total_distance:.1f} km', '#0066cc'),
                ('총 산출금액', f'{total_amount:,}원', '#0071e3'),
                ('불일치', f'{n_mis}건', '#d70015'),
            ]
            for idx, (label, value, color) in enumerate(summaries):
                col = idx % 2
                row = idx // 2
                item_frame = tk.Frame(summary_grid, bg='white', relief=tk.RAISED, bd=1)
                item_frame.grid(row=row, column=col, padx=10, pady=5, sticky='ew')
                tk.Label(item_frame, text=label, font=('맑은 고딕', 9), bg='white', fg='#6e6e73').pack(pady=(8, 2), padx=20)
                tk.Label(item_frame, text=value, font=('맑은 고딕', 12, 'bold'), bg='white', fg=color).pack(pady=(2, 8), padx=20)
        table_label_frame = tk.LabelFrame(result_frame, text='상세 내역', font=('맑은 고딕', 11, 'bold'), bg='white', fg='#1d1d1f', padx=10, pady=10)
        table_label_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(10, 20))
        filter_frame = tk.Frame(table_label_frame, bg='white')
        filter_frame.pack(fill=tk.X, pady=(0, 8))
        tk.Label(filter_frame, text='차종:', font=('맑은 고딕', 10), bg='white').pack(side=tk.LEFT, padx=(0, 5))
        vehicle_filter = ttk.Combobox(filter_frame, values=['전체', '오토바이', '다마스', '1t', '연계'], font=('맑은 고딕', 10), state='readonly', width=10)
        vehicle_filter.current(0)
        vehicle_filter.pack(side=tk.LEFT, padx=(0, 15))
        tk.Label(filter_frame, text='편도/왕복:', font=('맑은 고딕', 10), bg='white').pack(side=tk.LEFT, padx=(0, 5))
        trip_filter = ttk.Combobox(filter_frame, values=['전체', '편도', '왕복'], font=('맑은 고딕', 10), state='readonly', width=10)
        trip_filter.current(0)
        trip_filter.pack(side=tk.LEFT, padx=(0, 15))
        filter_count_label = tk.Label(filter_frame, text='', font=('맑은 고딕', 10), bg='white', fg='#6e6e73')
        filter_count_label.pack(side=tk.LEFT)
        table_container = tk.Frame(table_label_frame, bg='white')
        table_container.pack(fill=tk.BOTH, expand=True)
        scrollbar_y = tk.Scrollbar(table_container, orient=tk.VERTICAL)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x = tk.Scrollbar(table_container, orient=tk.HORIZONTAL)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        display_columns = ['NO', '출발지', '도착지', '계산거리(km)', '비용', '단가표금액', '차액', '검증결과', '적용기준']
        cols_in = list(self.result_df.columns) if self.result_df is not None else []
        available_columns = [c for c in display_columns if c in cols_in or c == 'NO']
        tree = ttk.Treeview(table_container, columns=available_columns, show='headings', yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set, height=12)
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)
        for col in available_columns:
            if col == 'NO':
                tree.column(col, width=50, anchor='center')
            elif col in ('출발지', '도착지'):
                tree.column(col, width=200, anchor='w')
            elif col == '계산거리(km)':
                tree.column(col, width=110, anchor='e')
            elif col in ('비용', '단가표금액', '차액'):
                tree.column(col, width=100, anchor='e')
            elif col == '적용기준':
                tree.column(col, width=260, anchor='w')
            else:
                tree.column(col, width=110, anchor='center')
            tree.heading(col, text=col, anchor='center')
        def refresh_tree(*args):
            vf = vehicle_filter.get()
            tf = trip_filter.get()
            for item in tree.get_children():
                tree.delete(item)
            shown = 0
            if self.result_df is not None:
                for n, (idx, row) in enumerate(self.result_df.iterrows()):
                    rv = str(row.get('차종', '')) if '차종' in row.index else ''
                    rt = str(row.get('편도/왕복', '')) if '편도/왕복' in row.index else ''
                    if vf != '전체' and rv != vf:
                        continue
                    if tf != '전체' and rt != tf:
                        continue
                    values = []
                    for col in available_columns:
                        if col == 'NO':
                            values.append(n + 1)
                            continue
                        if col not in row.index:
                            values.append('')
                            continue
                        value = row[col]
                        try:
                            is_val = bool(pd.notna(value))
                        except Exception:
                            is_val = True
                        if col in ('비용', '단가표금액', '차액') and is_val:
                            try:
                                values.append(f'{int(value):,}')
                            except Exception:
                                values.append(str(value))
                            continue
                        if col == '계산거리(km)' and is_val:
                            if isinstance(value, (int, float)):
                                values.append(f'{value:.1f}')
                            else:
                                values.append(str(value))
                            continue
                        values.append(str(value) if is_val else '')
                    tag = 'mismatch' if row.get('검증결과') == '불일치' else ''
                    tree.insert('', tk.END, values=values, tags=(tag,))
                    shown += 1
            tree.tag_configure('mismatch', background='#ffe6e6')
            total = len(self.result_df) if self.result_df is not None else 0
            filter_count_label.config(text=f'{shown}/{total}건 표시')
        vehicle_filter.bind('<<ComboboxSelected>>', refresh_tree)
        trip_filter.bind('<<ComboboxSelected>>', refresh_tree)
        refresh_tree()
        tree.pack(fill=tk.BOTH, expand=True)
        button_frame = tk.Frame(self.main_container, bg='#f5f5f7')
        button_frame.pack(fill=tk.X, pady=20)
        btn_container = tk.Frame(button_frame, bg='#f5f5f7')
        btn_container.pack()
        tk.Button(btn_container, text='◀ 이전 단계', font=('맑은 고딕', 12, 'bold'), bg='#6e6e73', fg='white', padx=30, pady=12, cursor='hand2', relief=tk.RAISED, bd=2, command=self.show_step1_screen).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_container, text='💾 결과 저장', font=('맑은 고딕', 12, 'bold'), bg='#0071e3', fg='white', padx=30, pady=12, cursor='hand2', relief=tk.RAISED, bd=2, command=self.save_result).pack(side=tk.LEFT, padx=10)

    __static_attributes__ = ('category_combo', 'current_screen', 'file_display_area', 'file_hint_shown', 'input_df', 'input_file_path', 'main_container', 'pricing_area_df', 'pricing_distance_df', 'pricing_file_path', 'result_df', 'root', 'section2', 'tonnage_combo', 'trip_type_combo', 'vehicle_combo', 'vehicle_container', 'vehicle_type_combo')


def main():
    root = tk.Tk()
    app = LogisticsSystem(root)
    root.mainloop()

if __name__ == '__main__':
    main()
