import os
import time
import glob
import re
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
dotenv_path = os.path.join(current_dir, '..', '.env')
load_dotenv(dotenv_path)

class ZeniusAutomation:
    def __init__(self):
        print(f"🔄 [{datetime.now().strftime('%H:%M:%S')}] Zenius 모니터링 시작...")
        
        self.config = {
            "CHROME_BIN": os.getenv("CHROME_BIN_PATH"),
            "DRIVER_PATH": os.path.join(os.getenv("CHROME_DRIVER_PATH"), "chromedriver.exe"),
            "URL": os.getenv("ZENIUS_URL"),
            "USER_ID": os.getenv("ZENIUS_USER_ID"),
            "USER_PW": os.getenv("ZENIUS_USER_PW"),
            "OUTPUT_PATH": os.getenv("OUTPUT_FILE_PATH")
        }
        
        self.tree_path = [i.strip() for i in os.getenv("TREE_PATH", "").split(",") if i]
        self.targets = [i.strip() for i in os.getenv("TEST_TARGET_SERVERS", "").split(",") if i]
        
        self.driver = self._setup_driver()
        self.wait = WebDriverWait(self.driver, 20)
        self.results = []
        self.prev_fs = {} # 전일 사용량 저장용
        self.re_num = re.compile(r'[^0-9.]')

    def _setup_driver(self):
        opts = Options()
        if self.config["CHROME_BIN"]: opts.binary_location = self.config["CHROME_BIN"]
        opts.add_argument('--headless')
        opts.add_argument('--no-sandbox')
        opts.add_argument('--disable-dev-shm-usage')
        opts.add_argument('--ignore-certificate-errors')
        return webdriver.Chrome(service=Service(self.config["DRIVER_PATH"]), options=opts)

    def load_previous_data(self):
        """전날 엑셀에서 마운트 경로별 사용량을 읽어옴"""
        files = glob.glob(os.path.join(self.config["OUTPUT_PATH"], "HDI Unix 서버 모니터링 정보_*.xlsx"))
        if not files: return
        try:
            latest = max(files, key=os.path.getctime)
            print(f"📂 이전 데이터 참조: {os.path.basename(latest)}")
            with pd.ExcelFile(latest) as xls:
                for sn in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sn)
                    # 데이터 시작 지점(7행)부터 마운트경로와 사용량 추출
                    fs_df = df.iloc[6:].copy()
                    fs_df.columns = fs_df.iloc[0]
                    fs_df = fs_df[1:]
                    if "마운트경로" in fs_df.columns and "사용량" in fs_df.columns:
                        self.prev_fs[sn.lower()] = dict(zip(fs_df["마운트경로"], fs_df["사용량"]))
        except Exception as e:
            print(f"⚠️ 이전 데이터 로드 실패: {e}")

    def _calculate_diff(self, curr, prev):
        """숫자 차이 계산 (사용량 기준)"""
        try:
            c = float(self.re_num.sub('', str(curr)) or 0)
            p = float(self.re_num.sub('', str(prev)) or 0)
            diff = round(c - p, 2)
            
            unit = "GB" if "GB" in str(curr) else ("MB" if "MB" in str(curr) else "")
            if diff > 0: return f"▲ {diff}{unit}"
            elif diff < 0: return f"▼ {abs(diff)}{unit}"
            return "(-)"
        except: return "(-)"

    def login(self):
        try:
            print(f"🔐 로그인 중...")
            self.driver.get(self.config["URL"])
            self.wait.until(EC.presence_of_element_located((By.ID, "z_accountname"))).send_keys(self.config["USER_ID"])
            self.driver.find_element(By.ID, "z_accountsecret").send_keys(self.config["USER_PW"])
            self.driver.find_element(By.ID, "loginBtn").click()
            
            self.wait.until(EC.presence_of_element_located((By.CLASS_NAME, "treeclick")))
            for folder in self.tree_path:
                xpath = f"//tr[descendant::*[contains(text(), '{folder}')]]"
                row = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                icon = row.find_element(By.XPATH, ".//div[contains(@class, 'treeclick')]")
                if "tree-minus" not in icon.get_attribute("class"):
                    icon.click()
                    time.sleep(1.5)
            print("🌳 트리 메뉴 구성 완료")
        except Exception as e:
            print(f"❌ 로그인 오류: {e}"); self.driver.quit(); exit(1)

    def collect_data(self):
        main_win = self.driver.current_window_handle
        for server in self.targets:
            print(f"🚀 {server} 분석 중...")
            try:
                self.driver.switch_to.window(main_win)
                server_link = self.wait.until(EC.element_to_be_clickable((By.LINK_TEXT, server)))
                server_link.click()
                
                self.wait.until(lambda d: len(d.window_handles) > 1)
                self.driver.switch_to.window(self.driver.window_handles[-1])
                
                # 기본 정보
                cpu = self.wait.until(EC.presence_of_element_located((By.ID, 'cpuArea_Utilization_str'))).text.strip()
                phys = self.driver.find_element(By.CSS_SELECTOR, '#phyUsedGauge .half_grp_num').text.strip()
                swap = self.driver.find_element(By.CSS_SELECTOR, '#swapUsedGauge .half_grp_num').text.strip()
                
                # 디스크 탭 클릭 및 데이터 로딩 대기
                self.driver.find_element(By.ID, 'Disk').click()
                time.sleep(3) # 데이터 로딩을 위해 넉넉히 대기
                
                rows = self.driver.find_elements(By.CSS_SELECTOR, "tr.ui-widget-content.jqgrow")
                for row in rows:
                    cols = row.find_elements(By.TAG_NAME, "td")
                    if len(cols) > 9:
                        self.results.append({
                            "서버이름": server, "CPU": cpu, "Phys": phys, "Swap": swap,
                            "FS": cols[2].text.strip(), "경로": cols[3].text.strip(),
                            "전체": cols[6].text.strip(), "사용": cols[7].text.strip(),
                            "가용": cols[8].text.strip(), "율": cols[9].text.strip()
                        })
                print(f"  ✅ {server} 수집 완료")
                self.driver.close()
            except Exception as e:
                print(f"  ⚠️ {server} 오류: {e}")
                if len(self.driver.window_handles) > 1: self.driver.close()

    def save_report(self):
        if not self.results: return
        df = pd.DataFrame(self.results)
        path = os.path.join(self.config["OUTPUT_PATH"], f"[TEST]HDI Unix 서버 모니터링 정보_{datetime.now().strftime('%Y%m%d')}.xlsx")
        
        with pd.ExcelWriter(path, engine='openpyxl') as writer:
            for server in self.targets:
                sn = server[:31]; s_key = server.lower()
                group = df[df["서버이름"] == server]
                if group.empty: continue

                # 상단 요약 (CPU/Mem)
                summary = [{"항목": "CPU 사용률", "현재 수치": group.iloc[0]["CPU"], "전일 대비": "(-)"},
                           {"항목": "Physical Memory", "현재 수치": group.iloc[0]["Phys"], "전일 대비": "(-)"},
                           {"항목": "Swap Memory", "현재 수치": group.iloc[0]["Swap"], "전일 대비": "(-)"}]
                pd.DataFrame(summary).to_excel(writer, sheet_name=sn, index=False)

                # 파일시스템 테이블 (증감 로직 포함)
                fs_table = []
                prev_data = self.prev_fs.get(s_key, {})
                for _, r in group.iterrows():
                    curr_u = r["사용"]
                    prev_u = prev_data.get(r["경로"], "N/A")
                    fs_table.append({
                        "파일시스템": r["FS"], "마운트경로": r["경로"], "전체용량": r["전체"],
                        "사용량": curr_u, "사용률(현재)": r["율"],
                        "전일 대비 증감": self._calculate_diff(curr_u, prev_u)
                    })
                
                pd.DataFrame(fs_table).to_excel(writer, sheet_name=sn, index=False, startrow=6)
        print(f"✨ 리포트 생성 완료: {path}")

    def run(self):
        try:
            self.load_previous_data()
            self.login()
            self.collect_data()
            self.save_report()
        finally:
            self.driver.quit()

if __name__ == "__main__":
    ZeniusAutomation().run()