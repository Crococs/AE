import os
import time
import json
import glob
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl.styles import PatternFill, Font, Alignment

# --- [환경 설정 로드] ---
current_dir = os.path.dirname(os.path.abspath(__file__))
# 상위 폴더의 .env 파일을 가리킴
dotenv_path = os.path.join(current_dir, '..', '.env')
load_dotenv(dotenv_path)

class ZeniusAutomation:
    def __init__(self):
        print(f"🔄 [{datetime.now().strftime('%H:%M:%S')}] Zenius 자동화 초기화...")
        
        # 환경 변수 로드 확인
        driver_dir = os.getenv("CHROME_DRIVER_PATH")
        if driver_dir is None:
            print("❌ 오류: .env 파일에서 'CHROME_DRIVER_PATH'를 읽지 못했습니다.")
            print(f"현재 참조 중인 .env 경로: {os.path.abspath(dotenv_path)}")
            exit(1)

        self.config = {
            "CHROME_BIN": os.getenv("CHROME_BIN_PATH"),
            "DRIVER_PATH": os.path.join(driver_dir, "chromedriver.exe"),
            "URL": os.getenv("ZENIUS_URL"),
            "USER_ID": os.getenv("ZENIUS_USER_ID"),
            "USER_PW": os.getenv("ZENIUS_USER_PW"),
            "OUTPUT_PATH": os.getenv("OUTPUT_FILE_PATH")
        }
        
        self.tree_path = [item.strip() for item in os.getenv("TREE_PATH", "").split(",") if item]
        self.targets = [item.strip() for item in os.getenv("TARGET_SERVERS", "").split(",") if item]
        
        self.driver = self._setup_driver()
        self.wait = WebDriverWait(self.driver, 15)
        self.results = []
        self.prev_perf = {}
        self.prev_fs = {}

    def _setup_driver(self):
        options = Options()
        if self.config["CHROME_BIN"]: 
            options.binary_location = self.config["CHROME_BIN"]
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--ignore-certificate-errors')
        options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
        
        service = Service(self.config["DRIVER_PATH"])
        return webdriver.Chrome(service=service, options=options)

    def login(self):
        try:
            print(f"🔐 Zenius 접속 및 로그인 중...")
            self.driver.get(self.config["URL"])
            self.wait.until(EC.presence_of_element_located((By.ID, "z_accountname"))).send_keys(self.config["USER_ID"])
            self.driver.find_element(By.ID, "z_accountsecret").send_keys(self.config["USER_PW"])
            self.driver.find_element(By.ID, "loginBtn").click()
            self.wait.until(EC.presence_of_element_located((By.CLASS_NAME, "treeclick")))
            
            for folder in self.tree_path:
                row_xpath = f"//tr[descendant::*[contains(text(), '{folder}')]]"
                target_row = self.wait.until(EC.presence_of_element_located((By.XPATH, row_xpath)))
                tree_icon = target_row.find_element(By.XPATH, ".//div[contains(@class, 'treeclick')]")
                if "tree-minus" not in tree_icon.get_attribute("class"):
                    self.driver.execute_script("arguments[0].click();", tree_icon)
                    time.sleep(1)
        except Exception as e:
            print(f"❌ 로그인 실패: {e}"); self.driver.quit(); exit(1)

    def collect_data(self):
        main_window = self.driver.current_window_handle
        for server in self.targets:
            print(f"🚀 {server} 추출 중...", end="\r")
            try:
                self.driver.switch_to.window(main_window)
                click_js = f"var t='{server}'; var a=Array.from(document.querySelectorAll('a')).find(e=>e.textContent.trim().toLowerCase()===t.toLowerCase()); if(a){{a.click(); return true;}} return false;"
                if not self.driver.execute_script(click_js): continue
                
                self.wait.until(lambda d: len(d.window_handles) > 1)
                self.driver.switch_to.window(self.driver.window_handles[-1])
                self.wait.until(EC.presence_of_element_located((By.ID, 'cpuArea_Utilization_str')))

                perf_raw = json.loads(self.driver.execute_script("""
                    let p = {
                        cpu: document.getElementById('cpuArea_Utilization_str')?.textContent.trim() || 'N/A',
                        phys: document.querySelector('#phyUsedGauge .half_grp_num')?.textContent.trim() || 'N/A',
                        swap: document.querySelector('#swapUsedGauge .half_grp_num')?.textContent.trim() || 'N/A'
                    };
                    let d = document.querySelector('#Disk a') || document.querySelector('li#Disk');
                    if(d) d.click();
                    return JSON.stringify(p);
                """))
                time.sleep(2)
                fs_list = json.loads(self.driver.execute_script("""
                    let r = [];
                    document.querySelectorAll("tr.ui-widget-content.jqgrow").forEach(row => {
                        let c = row.querySelectorAll("td");
                        if(c.length > 9) r.push({
                            mount: c[3].textContent.trim(), pct: c[9].textContent.trim(),
                            total: c[6].textContent.trim(), used: c[7].textContent.trim(),
                            avail: c[8].textContent.trim(), fs: c[2].textContent.trim()
                        });
                    });
                    return JSON.stringify(r);
                """))

                for fs in fs_list:
                    self.results.append({
                        "서버이름": server, "CPU": perf_raw['cpu'], "Phys": perf_raw['phys'], "Swap": perf_raw['swap'],
                        "FS": fs['fs'], "경로": fs['mount'], "전체": fs['total'], "사용": fs['used'], "가용": fs['avail'], "율": fs['pct']
                    })
                self.driver.close()
            except: 
                if len(self.driver.window_handles) > 1: self.driver.close()

    def load_previous_data(self):
        files = glob.glob(os.path.join(self.config["OUTPUT_PATH"], "HDI Unix 서버 모니터링 정보_*.xlsx"))
        if not files: return
        try:
            latest = max(files, key=os.path.getctime)
            xls = pd.ExcelFile(latest)
            for sn in xls.sheet_names:
                df = pd.read_excel(latest, sheet_name=sn)
                s_key = sn.lower()
                self.prev_perf[s_key] = dict(zip(df.iloc[0:3]["항목"], df.iloc[0:3]["현재 수치"]))
                fs_df = df.iloc[6:].copy()
                fs_df.columns = fs_df.iloc[0]; fs_df = fs_df[1:]
                if "마운트경로" in fs_df.columns:
                    self.prev_fs[s_key] = dict(zip(fs_df["마운트경로"], fs_df["사용률(현재)"]))
        except: pass

    def _calculate_diff(self, curr, prev):
        try:
            c = float(str(curr).replace('%','').strip())
            p = float(str(prev).replace('%','').strip())
            diff = round(c - p, 2)
            if diff > 0: return f"▲ {diff}%"
            elif diff < 0: return f"▼ {abs(diff)}%"
            return "(-)"
        except: return "(-)"

    def save_report(self):
        if not self.results: return
        df_total = pd.DataFrame(self.results)
        output_file = os.path.join(self.config["OUTPUT_PATH"], f"HDI Unix 서버 모니터링 정보_{datetime.now().strftime('%Y%m%d')}.xlsx")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for target in self.targets:
                sn = str(target)[:31]; s_key = target.lower()
                group = df_total[df_total["서버이름"].str.lower() == s_key]
                if group.empty: continue

                summary = []
                for lbl, col in {"CPU 사용률":"CPU", "Physical Memory":"Phys", "Swap Memory":"Swap"}.items():
                    curr = group.iloc[0][col]
                    prev = self.prev_perf.get(s_key, {}).get(lbl, "N/A")
                    summary.append({"항목": lbl, "현재 수치": curr, "전일 대비": self._calculate_diff(curr, prev)})
                pd.DataFrame(summary).to_excel(writer, sheet_name=sn, index=False, startrow=0)

                fs_table = []
                for _, r in group.iterrows():
                    prev_pct = self.prev_fs.get(s_key, {}).get(r["경로"], "N/A")
                    fs_table.append({
                        "파일시스템": r["FS"], "마운트경로": r["경로"], "전체용량": r["전체"], "사용량": r["사용"], 
                        "사용률(현재)": r["율"], "전일 대비 증감": self._calculate_diff(r["율"], prev_pct)
                    })
                fs_table.sort(key=lambda x: (-1 if 'all' in str(x["파일시스템"]).lower() else 0, -float(str(x["사용률(현재)"]).replace('%','').strip() or 0)))
                pd.DataFrame(fs_table).to_excel(writer, sheet_name=sn, index=False, startrow=6)
                self._apply_styling(writer.book[sn])
        print(f"\n✨ 리포트 생성 완료: {output_file}")

    def _apply_styling(self, ws):
        """배경색 늘어짐 방지 스타일링"""
        header_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        no_fill = PatternFill(fill_type=None)
        bold_font = Font(bold=True)
        red_font = Font(color="9C0006", bold=True)
        center_ali = Alignment(horizontal='center', vertical='center')

        # 전체 초기화
        for row in ws.iter_rows():
            for cell in row:
                cell.fill = no_fill
                cell.alignment = center_ali

        # 상단 표 (A~C열만)
        for row_idx in range(1, 5):
            for col_idx in range(1, 4):
                cell = ws.cell(row=row_idx, column=col_idx)
                if row_idx == 1:
                    cell.fill = header_fill
                    cell.font = bold_font

        # 하단 표 (A~F열만)
        for row_idx in range(7, ws.max_row + 1):
            for col_idx in range(1, 7):
                cell = ws.cell(row=row_idx, column=col_idx)
                if row_idx == 7:
                    cell.fill = header_fill
                    cell.font = bold_font
                if row_idx > 7 and col_idx == 5: # 사용률 강조
                    try:
                        val = float(str(cell.value).replace('%','').strip())
                        if val >= 90:
                            cell.fill = red_fill
                            cell.font = red_font
                    except: pass

    def run(self):
        try:
            self.load_previous_data()
            self.login()
            self.collect_data()
            self.save_report()
        finally:
            if hasattr(self, 'driver'):
                self.driver.quit()

if __name__ == "__main__":
    ZeniusAutomation().run()