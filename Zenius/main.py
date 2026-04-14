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
        self.targets = [i.strip() for i in os.getenv("TARGET_SERVERS", "").split(",") if i]
        
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
        # 최신 파일 탐색
        files = [f for f in glob.glob(os.path.join(self.config["OUTPUT_PATH"], "HDI Unix 서버 모니터링 정보_*.xlsx")) 
                 if "~$" not in f]
        if not files: 
            print("ℹ️ 참조할 이전 데이터 파일이 없습니다.")
            self.prev_summary = {}
            return
            
        try:
            # 파일명 날짜 기준 정렬 (최신 파일 선택)
            latest = sorted(files, key=lambda x: re.findall(r'\d{8}', x), reverse=True)[0]
            print(f"📂 가장 최근 데이터 참조: {os.path.basename(latest)}")
            
            self.prev_summary = {}
            self.prev_fs = {}

            with pd.ExcelFile(latest) as xls:
                for sn in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sn, header=None)
                    s_key = sn.strip().lower()

                    # --- [상단 요약 정보 추출] 1~4행 부근 ---
                    # 엑셀의 0~5행 사이에서 CPU, Physical, Swap 키워드가 있는 행을 찾아 데이터 매핑
                    temp_summary = {}
                    for i in range(min(len(df), 6)):
                        row_vals = [str(v).strip() for v in df.iloc[i].values]
                        if "CPU 사용률" in row_vals[0]: temp_summary["CPU"] = row_vals[1]
                        elif "Physical Memory" in row_vals[0]: temp_summary["Phys"] = row_vals[1]
                        elif "Swap Memory" in row_vals[0]: temp_summary["Swap"] = row_vals[1]
                    
                    if temp_summary:
                        self.prev_summary[s_key] = temp_summary

                    # --- [파일시스템 정보 추출] ---
                    header_row_idx = -1
                    for i in range(len(df)):
                        if "마운트경로" in df.iloc[i].values:
                            header_row_idx = i
                            break
                    
                    if header_row_idx != -1:
                        fs_df = df.iloc[header_row_idx:].copy()
                        fs_df.columns = fs_df.iloc[0]
                        fs_df = fs_df[1:]
                        
                        self.prev_fs[s_key] = {
                            str(k).strip(): str(v).strip() 
                            for k, v in zip(fs_df["마운트경로"], fs_df["사용량"]) 
                            if pd.notna(k)
                        }
        except Exception as e:
            print(f"⚠️ 이전 데이터 로드 실패: {e}")
            self.prev_summary = {}
            self.prev_fs = {}

    def _calculate_diff(self, curr, prev):
        """전일 사용량과 당일 사용량의 순수 차이값만 반환"""
        try:
            # 단위(GB, MB 등)를 제외한 숫자만 추출하여 계산
            c_val = float(self.re_num.sub('', str(curr)) or 0)
            p_val = float(self.re_num.sub('', str(prev)) or 0)
            
            diff = round(c_val - p_val, 2)
            
            # 원본 데이터에서 단위 추출 (없으면 빈 문자열)
            unit = ""
            if "GB" in str(curr): unit = "GB"
            elif "MB" in str(curr): unit = "MB"
            
            if diff > 0:
                return f"+{diff}{unit}"
            elif diff < 0:
                # abs()를 써서 -가 중복되지 않게 처리 (예: -0.5GB)
                return f"{diff}{unit}" 
            else:
                return "0" # 변동 없음
        except:
            return "-" # 비교 불가 시

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
                
                # --- 수정된 부분: lambda 대신 공식 EC 사용 ---
                self.wait.until(EC.number_of_windows_to_be(2))
                time.sleep(0.5) # 창 전환 직후 안정화를 위한 짧은 휴식
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
        if not self.results: 
            print("⚠️ 수집된 데이터가 없어 리포트를 생성하지 않습니다.")
            return
            
        df = pd.DataFrame(self.results)
        # 오늘 날짜 파일명 생성
        base_name = f"HDI Unix 서버 모니터링 정보_{datetime.now().strftime('%Y%m%d')}"
        path = os.path.join(self.config["OUTPUT_PATH"], f"{base_name}.xlsx")
        
        # 파일 열림 에러(PermissionError) 방지
        try:
            writer = pd.ExcelWriter(path, engine='openpyxl')
        except PermissionError:
            path = os.path.join(self.config["OUTPUT_PATH"], f"{base_name}_{datetime.now().strftime('%H%M%S')}.xlsx")
            print(f"⚠️ 기존 파일이 열려 있어 새 이름으로 저장합니다: {os.path.basename(path)}")
            writer = pd.ExcelWriter(path, engine='openpyxl')

        with writer:
            for server in self.targets:
                sn = server[:31]
                s_key = server.strip().lower()
                
                group = df[df["서버이름"] == server]
                if group.empty: continue

                # 현재 데이터 추출
                curr_cpu = str(group["CPU"].values[0])
                curr_phys = str(group["Phys"].values[0])
                curr_swap = str(group["Swap"].values[0])

                # 이전 요약 데이터와 비교
                p_sum = self.prev_summary.get(s_key, {})
                diff_cpu = self._calculate_diff(curr_cpu, p_sum.get("CPU", "0"))
                diff_phys = self._calculate_diff(curr_phys, p_sum.get("Phys", "0"))
                diff_swap = self._calculate_diff(curr_swap, p_sum.get("Swap", "0"))

                # 상단 요약 테이블 구성
                summary_data = [
                    {"항목": "CPU 사용률", "현재 수치": curr_cpu, "전일 대비": diff_cpu},
                    {"항목": "Physical Memory", "현재 수치": curr_phys, "전일 대비": diff_phys},
                    {"항목": "Swap Memory", "현재 수치": curr_swap, "전일 대비": diff_swap}
                ]
                pd.DataFrame(summary_data).to_excel(writer, sheet_name=sn, index=False)

                # 파일시스템 테이블 구성
                fs_table = []
                prev_data = self.prev_fs.get(s_key, {})
                
                for _, r in group.iterrows():
                    curr_u = str(r["사용"])
                    curr_path = str(r["경로"]).strip()
                    prev_u = prev_data.get(curr_path, "N/A")
                    
                    fs_table.append({
                        "파일시스템": r["FS"], 
                        "마운트경로": r["경로"], 
                        "전체용량": r["전체"],
                        "사용량": curr_u, 
                        "사용률(현재)": r["율"],
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