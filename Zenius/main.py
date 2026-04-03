import os
import time
import json
import glob
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# .env 파일의 환경변수를 불러옴
current_dir = os.path.dirname(os.path.abspath(__file__))
dotenv_path = os.path.join(current_dir, '..', '.env')
if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)
    print("환경 변수 로드 성공!")
else:
    print(f"오류: {dotenv_path} 위치에 .env 파일이 없습니다.")

# --- [1. 사용자 설정] ---
CONFIG = {
    "PATH": os.getenv("CHROME_DRIVER_PATH"),
    "Output_file_Path": os.getenv("OUTPUT_FILE_PATH"),
    "CHROME_BIN": os.getenv("CHROME_BIN_PATH"),
    "ZENIUS_URL": os.getenv("ZENIUS_URL"),
    "USER_ID": os.getenv("ZENIUS_USER_ID"),
    "USER_PW": os.getenv("ZENIUS_USER_PW")
}
               
def get_env_list(key, default=""):
    value = os.getenv(key, default)
    if not value:
        return []
    # 공백이 섞여 있을 경우를 대비해 strip() 처리 포함
    return [item.strip() for item in value.split(",")]

TREE_PATH = get_env_list("TREE_PATH")
TARGET_SERVERS = get_env_list("TARGET_SERVERS")

# --- [2. 헬퍼 함수] ---
def calculate_diff(curr, prev):
    """현재 수치와 전일 수치를 비교하여 증감 표시"""
    try:
        c_val = float(str(curr).replace('%','').strip())
        p_val = float(str(prev).replace('%','').strip())
        diff = round(c_val - p_val, 2)
        if diff > 0: return f"▲ {diff}%"
        elif diff < 0: return f"▼ {abs(diff)}%"
        return "(-)"
    except: return "(-)"

# --- [3. 메인 로직] ---
def main():
    options = Options()
    options.binary_location = CONFIG["CHROME_BIN"]
    options.add_argument('--ignore-certificate-errors')
    
    driver_path = os.path.join(CONFIG["PATH"], "chromedriver.exe")
    driver = webdriver.Chrome(service=Service(driver_path), options=options)
    wait = WebDriverWait(driver, 15)
    final_results = []

    try:
        # 1. 로그인 및 트리 확장
        driver.get(CONFIG["ZENIUS_URL"])
        print("🔐 Zenius 로그인 중...")
        wait.until(EC.presence_of_element_located((By.ID, "z_accountname"))).send_keys(CONFIG["USER_ID"])
        driver.find_element(By.ID, "z_accountsecret").send_keys(CONFIG["USER_PW"])
        driver.find_element(By.ID, "loginBtn").click()
        time.sleep(5)
        main_window = driver.current_window_handle

        for folder in TREE_PATH:
            try:
                row_xpath = f"//tr[descendant::*[contains(text(), '{folder}')]]"
                target_row = wait.until(EC.presence_of_element_located((By.XPATH, row_xpath)))
                tree_icon = target_row.find_element(By.XPATH, ".//div[contains(@class, 'treeclick')]")
                if "tree-minus" not in tree_icon.get_attribute("class"):
                    driver.execute_script("arguments[0].click();", tree_icon)
                    time.sleep(2)
            except: pass

        # 2. 서버별 데이터 수집
        for server in TARGET_SERVERS:
            print(f"🚀 {server} 데이터 수집 중...")
            try:
                driver.switch_to.window(main_window)
                click_js = f"var t='{server}'; var a=Array.from(document.querySelectorAll('a')).find(e=>e.textContent.trim().toLowerCase()===t.toLowerCase()); if(a){{a.click(); return true;}} return false;"
                if not driver.execute_script(click_js): continue
                
                time.sleep(8)
                driver.switch_to.window(driver.window_handles[-1])

                # 성능 및 파일시스템 데이터 추출
                data_js = """
                    var d={p:{},f:[]};
                    var c=document.getElementById('cpuArea_Utilization_str'), m=document.querySelector('#phyUsedGauge .half_grp_num'), s=document.querySelector('#swapUsedGauge .half_grp_num');
                    d.p={cpu:c?c.textContent.trim():'N/A', phys:m?m.textContent.trim():'N/A', swap:s?s.textContent.trim():'N/A'};
                    var t=document.querySelector('#Disk a')||document.querySelector('li#Disk'); if(t)t.click();
                    return JSON.stringify(d);
                """
                perf_raw = json.loads(driver.execute_script(data_js))
                time.sleep(3)
                
                fs_js = """
                    var r=[]; document.querySelectorAll("tr.ui-widget-content.jqgrow").forEach(row=>{
                        var c=row.querySelectorAll("td");
                        if(c.length>9) r.push({mount:c[3].textContent.trim(), pct:c[9].textContent.trim(), total:c[6].textContent.trim(), used:c[7].textContent.trim(), avail:c[8].textContent.trim(), fs:c[2].textContent.trim()});
                    }); return JSON.stringify(r);
                """
                fs_list = json.loads(driver.execute_script(fs_js))

                for fs in fs_list:
                    final_results.append({
                        "서버이름": server, "CPU": perf_raw['p']['cpu'], "Phys": perf_raw['p']['phys'], "Swap": perf_raw['p']['swap'],
                        "FS": fs['fs'], "경로": fs['mount'], "전체": fs['total'], "사용": fs['used'], "가용": fs['avail'], "율": fs['pct']
                    })
                driver.close()
                driver.switch_to.window(main_window)
                print(f"✅ {server} 완료")
            except Exception as e:
                print(f"❌ {server} 에러: {e}")
                if len(driver.window_handles) > 1: driver.close()
                driver.switch_to.window(main_window)

        # 3. 전일 데이터 로드 (비교용)
        prev_perf, prev_fs = {}, {}
        files = glob.glob(os.path.join(CONFIG["Output_file_Path"], "HDI Unix 서버 모니터링 정보_*.xlsx"))
        if files:
            try:
                latest = max(files, key=os.path.getctime)
                print(f"📂 전일 파일 대조 중: {os.path.basename(latest)}")
                xls = pd.ExcelFile(latest)
                for sn in xls.sheet_names:
                    df = pd.read_excel(latest, sheet_name=sn)
                    s_key = sn.lower()
                    prev_perf[s_key] = dict(zip(df.iloc[0:3]["항목"], df.iloc[0:3]["현재 수치"]))
                    fs_df = df.iloc[6:].copy()
                    fs_df.columns = fs_df.iloc[0]; fs_df = fs_df[1:]
                    if "마운트경로" in fs_df.columns:
                        prev_fs[s_key] = dict(zip(fs_df["마운트경로"], fs_df["사용률(현재)"]))
            except: pass

        # 4. 엑셀 저장
        if final_results:
            df_total = pd.DataFrame(final_results)
            output_file = os.path.join(CONFIG["Output_file_Path"], f"HDI Unix 서버 모니터링 정보_{time.strftime('%Y%m%d')}.xlsx")
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for target in TARGET_SERVERS:
                    sn = str(target)[:31]; s_key = target.lower()
                    group = df_total[df_total["서버이름"].str.lower() == s_key]
                    if group.empty: continue

                    # 상단 요약 표
                    summary = []
                    for lbl, col in {"CPU 사용률":"CPU", "Physical Memory":"Phys", "Swap Memory":"Swap"}.items():
                        curr = group.iloc[0][col]
                        prev = prev_perf.get(s_key, {}).get(lbl, "N/A")
                        summary.append({"항목": lbl, "현재 수치": curr, "전일 대비": calculate_diff(curr, prev)})
                    pd.DataFrame(summary).to_excel(writer, sheet_name=sn, index=False, startrow=0)

                    # 하단 파일시스템 상세 표
                    fs_table = []
                    for _, r in group.iterrows():
                        curr_pct = r["율"]
                        prev_pct = prev_fs.get(s_key, {}).get(r["경로"], "N/A")
                        fs_table.append({
                            "파일시스템": r["FS"], "마운트경로": r["경로"], "전체용량": r["전체"], "사용량": r["사용"], 
                            "사용률(현재)": curr_pct, "전일 대비 증감": calculate_diff(curr_pct, prev_pct)
                        })
                    
                    # --- [All 상단 고정 + 나머지 내림차순 정렬] ---
                    def sort_key(x):
                        # All은 무조건 가장 위(-1)
                        if str(x["파일시스템"]).strip().lower() == 'all':
                            return (-1, 0)
                        # 나머지는 사용률 숫자 변환 후 내림차순(마이너스 부호)
                        try:
                            val = float(str(x["사용률(현재)"]).replace('%', '').strip() or 0)
                        except:
                            val = 0
                        return (0, -val)

                    fs_table.sort(key=sort_key)
                    # ------------------------------------------

                    pd.DataFrame(fs_table).to_excel(writer, sheet_name=sn, index=False, startrow=6)

            print(f"\n📊 [완료] 리포트 생성 완료: {output_file}")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()