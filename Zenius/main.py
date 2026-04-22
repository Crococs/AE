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

from webdriver_manager.chrome import ChromeDriverManager


# --- 환경 설정 ---
current_dir = os.path.dirname(os.path.abspath(__file__))
dotenv_path = os.path.join(current_dir, '..', '.env')
load_dotenv(dotenv_path)


class ZeniusAutomation:
    def __init__(self):
        print(f"🔄 [{datetime.now().strftime('%H:%M:%S')}] Zenius 모니터링 시작...")

        self.config = {
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
        self.prev_summary = {}
        self.prev_fs = {}

        self.re_num = re.compile(r'[^0-9.]')

    def _setup_driver(self):
        opts = Options()
        opts.add_argument('--headless')
        opts.add_argument('--no-sandbox')
        opts.add_argument('--disable-dev-shm-usage')
        opts.add_argument('--ignore-certificate-errors')
        opts.add_argument('--disable-blink-features=AutomationControlled')
        opts.add_argument('--disable-gpu')

        return webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=opts
        )

    # --------------------------
    # 이전 데이터 로드
    # --------------------------
    def load_previous_data(self):
        files = [
            f for f in glob.glob(os.path.join(self.config["OUTPUT_PATH"], "HDI Unix 서버 모니터링 정보_*.xlsx"))
            if "~$" not in f
        ]

        if not files:
            print("ℹ️ 이전 데이터 없음")
            return

        latest = sorted(files, key=lambda x: re.findall(r'\d{8}', x), reverse=True)[0]
        print(f"📂 이전 파일: {os.path.basename(latest)}")

        try:
            with pd.ExcelFile(latest) as xls:
                for sn in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sn, header=None)
                    s_key = sn.strip().lower()

                    temp_summary = {}
                    for i in range(min(len(df), 6)):
                        row_vals = [str(v).strip() for v in df.iloc[i].values]
                        if "CPU 사용률" in row_vals[0]:
                            temp_summary["CPU"] = row_vals[1]
                        elif "Physical Memory" in row_vals[0]:
                            temp_summary["Phys"] = row_vals[1]
                        elif "Swap Memory" in row_vals[0]:
                            temp_summary["Swap"] = row_vals[1]

                    if temp_summary:
                        self.prev_summary[s_key] = temp_summary

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

    # --------------------------
    # 전일 대비 계산
    # --------------------------
    def _calculate_diff(self, curr, prev):
        try:
            c_val = float(self.re_num.sub('', str(curr)) or 0)
            p_val = float(self.re_num.sub('', str(prev)) or 0)
            diff = round(c_val - p_val, 2)

            unit = "GB" if "GB" in str(curr) else "MB" if "MB" in str(curr) else ""

            if diff > 0:
                return f"▲{diff}{unit}"
            elif diff < 0:
                return f"▽{abs(diff)}{unit}"
            else:
                return "-"
        except:
            return "-"

    # --------------------------
    # 로그인
    # --------------------------
    def login(self):
        print("🔐 로그인 중...")

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
                time.sleep(1)

        print("🌳 트리 완료")

    # --------------------------
    # 데이터 수집
    # --------------------------
    def collect_data(self):
        main_win = self.driver.current_window_handle

        for server in self.targets:
            print(f"🚀 {server} 분석 중...")

            for attempt in range(2):
                try:
                    self.driver.switch_to.window(main_win)

                    server_link = self.wait.until(
                        EC.element_to_be_clickable((By.LINK_TEXT, server))
                    )
                    server_link.click()

                    self.wait.until(EC.number_of_windows_to_be(2))
                    self.driver.switch_to.window(self.driver.window_handles[-1])

                    cpu = self.wait.until(
                        EC.presence_of_element_located((By.ID, 'cpuArea_Utilization_str'))
                    ).text.strip()

                    phys = self.wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, '#phyUsedGauge .half_grp_num'))
                    ).text.strip()

                    swap = self.wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, '#swapUsedGauge .half_grp_num'))
                    ).text.strip()

                    self.driver.find_element(By.ID, 'Disk').click()

                    self.wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "tr.ui-widget-content.jqgrow"))
                    )

                    rows = self.driver.find_elements(By.CSS_SELECTOR, "tr.ui-widget-content.jqgrow")

                    for row in rows:
                        cols = row.find_elements(By.TAG_NAME, "td")
                        if len(cols) > 9:
                            self.results.append({
                                "서버이름": server,
                                "CPU": cpu,
                                "Phys": phys,
                                "Swap": swap,
                                "FS": cols[2].text.strip(),
                                "경로": cols[3].text.strip(),
                                "전체": cols[6].text.strip(),
                                "사용": cols[7].text.strip(),
                                "가용": cols[8].text.strip(),
                                "율": cols[9].text.strip()
                            })

                    print(f"  ✅ {server} 완료")
                    self.driver.close()
                    break

                except Exception as e:
                    print(f"  ⚠️ {server} 오류 (시도 {attempt+1}): {e}")

                    if len(self.driver.window_handles) > 1:
                        self.driver.close()

                    if attempt == 1:
                        print(f"  ❌ {server} 실패")

    # --------------------------
    # 리포트 생성
    # --------------------------
    def save_report(self):
        if not self.results:
            print("⚠️ 데이터 없음")
            return

        df = pd.DataFrame(self.results)

        base_name = f"HDI Unix 서버 모니터링 정보_{datetime.now().strftime('%Y%m%d')}"
        path = os.path.join(self.config["OUTPUT_PATH"], f"{base_name}.xlsx")

        writer = pd.ExcelWriter(path, engine='openpyxl')

        with writer:
            for server in self.targets:
                sn = re.sub(r'[\\/*?:\[\]]', '', server)[:31]
                s_key = sn.strip().lower()

                group = df[df["서버이름"] == server]
                if group.empty:
                    continue

                # 상단 요약
                curr_cpu = str(group["CPU"].values[0])
                curr_phys = str(group["Phys"].values[0])
                curr_swap = str(group["Swap"].values[0])

                p_sum = self.prev_summary.get(s_key, {})

                summary_data = [
                    {"항목": "CPU 사용률", "현재 수치": curr_cpu, "전일 대비": self._calculate_diff(curr_cpu, p_sum.get("CPU", "0"))},
                    {"항목": "Physical Memory", "현재 수치": curr_phys, "전일 대비": self._calculate_diff(curr_phys, p_sum.get("Phys", "0"))},
                    {"항목": "Swap Memory", "현재 수치": curr_swap, "전일 대비": self._calculate_diff(curr_swap, p_sum.get("Swap", "0"))}
                ]

                pd.DataFrame(summary_data).to_excel(writer, sheet_name=sn, index=False)

                # 파일시스템
                prev_data = self.prev_fs.get(s_key, {})
                fs_table_raw = []

                for _, r in group.iterrows():
                    curr_u = str(r["사용"])
                    curr_path = str(r["경로"]).strip()
                    prev_u = prev_data.get(curr_path, "N/A")

                    try:
                        raw_usage = float(str(r["율"]).replace('%', '').strip())
                    except:
                        raw_usage = -1

                    fs_table_raw.append({
                        "파일시스템": r["FS"],
                        "마운트경로": curr_path,
                        "전체용량": r["전체"],
                        "사용량": curr_u,
                        "사용률(현재)": r["율"],
                        "전일 대비 증감": self._calculate_diff(curr_u, prev_u),
                        "_sort_val": raw_usage
                    })

                # ⭐ 정렬 로직 (최종 요구사항)
                def fs_sort_key(item):
                    path = str(item["마운트경로"]).strip().upper()
                    try:
                        usage = float(item["_sort_val"])
                    except:
                        usage = -1

                    if path == "ALL":
                        return (0, 0)
                    elif path == "/":
                        return (1, 0)
                    else:
                        return (2, -usage)

                fs_sorted = sorted(fs_table_raw, key=fs_sort_key)

                final_fs = [
                    {k: v for k, v in item.items() if k != "_sort_val"}
                    for item in fs_sorted
                ]

                pd.DataFrame(final_fs).to_excel(
                    writer,
                    sheet_name=sn,
                    index=False,
                    startrow=6
                )

        print(f"✨ 리포트 생성 완료: {path}")

    # --------------------------
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