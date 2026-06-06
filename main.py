"""
KAU 종합정보시스템 수강생 목록 → 엑셀 자동 수집

수백 명 규모로 직접 돌려보며 짚은 문제와, 그에 대한 설계 결정
  1. 로그인 자동화  : .env 의 id/pass 로 자동 로그인 (사람이 붙어 있을 필요 제거)
  2. 속도 대폭 개선 : 엑셀 저장을 '한 명마다' → '한 페이지마다'로 (직접 찾아낸 가장 큰 병목)
  3. 대기 최적화    : 고정 time.sleep 대신 조건을 만족하면 즉시 진행하는 WebDriverWait
  4. 페이지 로딩 전략 eager + 창 최대화

⚠️ '수강생 관리' 메뉴까지 들어가는 클릭 경로(AUTO_NAV_STEPS)는 사이트를 직접 봐야
   알 수 있어 비워뒀습니다. 비어 있으면 그 단계만 잠깐 수동(엔터)으로 진행합니다.
   채우는 방법은 파일 하단 주석 참고.
"""

import time
import re
import os
import subprocess

import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# ─────────────────────────────────────────────────────────
# 설정
# ─────────────────────────────────────────────────────────
load_dotenv()
USER_ID = os.getenv("id")
USER_PW = os.getenv("pass")


def _env_bool(key, default=False):
    return os.getenv(key, str(default)).strip().lower() in ("1", "true", "yes", "y", "on")


def _env_int(key, default):
    try:
        return int(os.getenv(key, default))
    except (TypeError, ValueError):
        return int(default)


HEADLESS = _env_bool("HEADLESS", False)

LOGIN_URL = "https://nportal.kau.ac.kr/webcrea/GB03/mdi/login.html"

# 로그인 칸을 자동 탐지하지 못할 때만 직접 지정 (CSS selector). 보통은 None 으로 둬도 됨.
LOGIN_ID_SEL = None   # 예) "#user_id"
LOGIN_PW_SEL = None   # 예) "#user_pw"

# 로그인 후 '수강생 관리' 화면까지 누르는 메뉴 순서(글자 그대로).
# 비우면([]) 그 단계만 수동(엔터)으로 진행.
MENU_PATH = ["일반행정", "부서업무", "안전교육원", "수강생관리"]

# 수집 기간 (.env 의 START_*/END_* 로 지정, 둘 다 포함)
START_YEAR = _env_int("START_YEAR", 2024)
START_MONTH = _env_int("START_MONTH", 11)
END_YEAR = _env_int("END_YEAR", 2026)
END_MONTH = _env_int("END_MONTH", 5)


def _build_periods(sy, sm, ey, em):
    periods, y, m = [], sy, sm
    while (y, m) <= (ey, em):
        periods.append((y, m))
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return periods


PERIODS = _build_periods(START_YEAR, START_MONTH, END_YEAR, END_MONTH)

DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
SAVE_PATH = os.path.join(DESKTOP, "KAU_MASTER_DATA.xlsx")
COLUMNS = ["No", "국문1", "국문2", "영문1", "영문2", "생년월일", "소속1", "소속2",
           "직책1", "직책2", "이메일1", "이메일2", "연락처1", "연락처2",
           "구분", "등급", "과정명", "교육기간", "수집연월"]

date_pattern = re.compile(r'\(\d{4}\.\d{2}\.\d{2}.*?\)')

# ─────────────────────────────────────────────────────────
# 기존 데이터 로드 (이어 달리기)
# ─────────────────────────────────────────────────────────
existing_data = []
new_data = []

if os.path.exists(SAVE_PATH):
    try:
        existing_data = pd.read_excel(SAVE_PATH).values.tolist()
        print(f"📦 기존 데이터 {len(existing_data)}건 확보")
    except Exception as e:
        print(f"⚠️ 기존 엑셀 읽기 실패, 새로 생성: {e}")
else:
    print("📢 기존 파일 없음 → 새로 생성")


def save_to_excel():
    """기존 + 신규 데이터를 합쳐 저장. (페이지 단위로만 호출 → 디스크 쓰기 최소화)"""
    try:
        df = pd.DataFrame(existing_data + new_data, columns=COLUMNS)
        df.to_excel(SAVE_PATH, index=False)
        print(f"    💾 저장: 기존 {len(existing_data)} + 신규 {len(new_data)} = {len(df)}건")
    except Exception as e:
        print(f"    ❌ 저장 실패: {e}")


# ─────────────────────────────────────────────────────────
# 드라이버
# ─────────────────────────────────────────────────────────
opts = webdriver.ChromeOptions()
opts.page_load_strategy = "eager"   # DOMContentLoaded 시점에 진행 (불필요한 대기 단축)
if HEADLESS:
    opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1920,1080")   # 좌표 파싱 일관성 위해 고정 크기
    opts.add_argument("--disable-gpu")
    print("🕶️ 헤드리스 모드로 실행")
else:
    opts.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
wait = WebDriverWait(driver, 20)


def auto_login():
    """.env 의 id/pass 로 자동 로그인. password 칸을 기준으로 폼을 자동 탐지."""
    if not USER_ID or not USER_PW:
        raise RuntimeError(".env 에서 id/pass 를 읽지 못했습니다.")

    driver.get(LOGIN_URL)

    # 1) 비밀번호 칸 찾기 (기본 문서 → 안 되면 iframe 순회)
    pw_field = None
    driver.switch_to.default_content()
    try:
        sel = LOGIN_PW_SEL or "input[type='password']"
        pw_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
    except TimeoutException:
        for fr in driver.find_elements(By.TAG_NAME, "iframe"):
            driver.switch_to.default_content()
            driver.switch_to.frame(fr)
            cand = driver.find_elements(By.CSS_SELECTOR, LOGIN_PW_SEL or "input[type='password']")
            if cand:
                pw_field = cand[0]
                break
    if not pw_field:
        raise RuntimeError("로그인 폼을 못 찾음 → LOGIN_ID_SEL/LOGIN_PW_SEL 을 직접 지정하세요.")

    # 2) 아이디 칸 찾기 (password 같은 컨텍스트의 보이는 text 입력칸)
    if LOGIN_ID_SEL:
        id_field = driver.find_element(By.CSS_SELECTOR, LOGIN_ID_SEL)
    else:
        id_field = None
        for el in driver.find_elements(By.CSS_SELECTOR, "input"):
            t = (el.get_attribute("type") or "text").lower()
            if t in ("text", "id", "email", "") and el.is_displayed():
                id_field = el
                break
        if id_field is None:
            raise RuntimeError("아이디 입력칸을 못 찾음 → LOGIN_ID_SEL 을 직접 지정하세요.")

    id_field.clear(); id_field.send_keys(USER_ID)
    pw_field.clear(); pw_field.send_keys(USER_PW)
    pw_field.send_keys(Keys.RETURN)
    print("✅ 자동 로그인 시도 완료")
    driver.switch_to.default_content()


def _try_click_in_current(label):
    """현재 프레임에서 label 글자를 가진 요소를 찾아 클릭. 성공 시 True."""
    xpaths = [
        f"//*[normalize-space(.)='{label}']",            # 정확히 일치(가장 안전)
        f"//*[contains(normalize-space(.), '{label}')]",  # 부분 일치(보조)
    ]
    for xp in xpaths:
        els = [e for e in driver.find_elements(By.XPATH, xp) if e.is_displayed()]
        if not els:
            continue
        els.sort(key=lambda e: len(e.text))  # 텍스트가 가장 짧은 = 가장 구체적인 요소
        el = els[0]
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            try:
                el.click()
            except Exception:
                driver.execute_script("arguments[0].click();", el)
            return True
        except Exception:
            continue
    return False


def _click_menu(label, timeout=15):
    """기본 문서 + 모든 iframe 을 뒤져 label 메뉴를 클릭. 트리 메뉴가 펼쳐질 때까지 재시도."""
    end_at = time.monotonic() + timeout
    while time.monotonic() < end_at:
        driver.switch_to.default_content()
        if _try_click_in_current(label):
            return True
        for fr in driver.find_elements(By.TAG_NAME, "iframe"):
            driver.switch_to.default_content()
            try:
                driver.switch_to.frame(fr)
            except Exception:
                continue
            if _try_click_in_current(label):
                return True
        time.sleep(0.3)
    return False


def navigate_to_roster():
    """MENU_PATH 순서대로 메뉴를 자동 클릭. 못 찾으면 그 항목만 수동 진행."""
    if not MENU_PATH:
        print("➡️  '수강생 관리' 화면으로 직접 이동한 뒤 엔터를 누르세요.")
        input()
    else:
        for label in MENU_PATH:
            if _click_menu(label):
                print(f"   ✓ '{label}' 클릭")
            else:
                print(f"⚠️ '{label}' 메뉴를 자동으로 못 찾았습니다.")
                input(f"   → 직접 '{label}' 클릭 후 엔터: ")
            time.sleep(0.6)  # 다음 하위 메뉴가 펼쳐질 시간
        print("✅ '수강생 관리' 화면 이동 완료")
    driver.switch_to.default_content()


def enter_work_frame():
    driver.switch_to.default_content()
    try:
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it("kas01_005_t"))
    except TimeoutException:
        driver.switch_to.frame("WorkFrame")


# ─────────────────────────────────────────────────────────
# 메인
# ─────────────────────────────────────────────────────────
main_window = None
try:
    auto_login()
    navigate_to_roster()
    enter_work_frame()

    for year, month in PERIODS:
        print(f"\n📅 [{year}년 {month:02d}월] 시작")
        driver.execute_script(
            f"document.querySelector('#SEARCH_FORM\\\\.SEARCH_HYEAR').value = '{year}';")
        driver.execute_script(
            f"document.querySelector('#SEARCH_FORM\\\\.SEARCH_MON').value = '{month:02d}';")
        driver.find_element(By.ID, "FormPush1.Btn_Select").click()

        # 검색 결과(과정 목록)가 뜰 때까지 대기 (고정 sleep(4) 대체)
        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, r"[id^='List1\.GWAJUNG_KNM']")))
        except TimeoutException:
            print("    📢 해당 월 과정 없음")
            continue
        time.sleep(0.5)  # 목록 렌더 안정화용 최소 대기

        course_elements = driver.find_elements(
            By.CSS_SELECTOR, r"[id^='List1\.GWAJUNG_KNM']")

        for i in range(len(course_elements)):
            try:
                course_main_text = course_elements[i].text.strip()
                driver.execute_script("arguments[0].click();", course_elements[i])

                # 인쇄 버튼이 클릭 가능해질 때까지 대기 (고정 sleep(1.5) 대체)
                prt = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "FormPush2.Btn_PRT2")))
                prt.click()

                WebDriverWait(driver, 15).until(EC.number_of_windows_to_be(2))
                main_window = driver.current_window_handle
                for handle in driver.window_handles:
                    if handle != main_window:
                        driver.switch_to.window(handle)
                        break

                # 리포트 뷰어 첫 렌더 대기 (고정 sleep(5) 대체)
                driver.switch_to.default_content()
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "#m2soft-crownix-text > div")))

                last_page_fingerprint = "INITIAL_EMPTY_VALUE"
                page_count = 1

                while True:
                    driver.switch_to.default_content()

                    # 페이지 전환 감지 (이전 페이지와 내용이 바뀔 때까지 폴링)
                    page_loaded = False
                    for _ in range(40):
                        report_divs = driver.find_elements(
                            By.CSS_SELECTOR, "#m2soft-crownix-text > div")
                        temp = [d.text.strip() for d in report_divs if d.text.strip()]
                        current_check = "".join(temp[:3]) if temp else ""
                        if (last_page_fingerprint == "INITIAL_EMPTY_VALUE"
                                or current_check != last_page_fingerprint):
                            page_loaded = True
                            break
                        time.sleep(0.2)

                    if not page_loaded:
                        print("    📢 마지막 페이지 완료")
                        break

                    report_divs = driver.find_elements(
                        By.CSS_SELECTOR, "#m2soft-crownix-text > div")
                    items, edu_period_val = [], ""
                    for d in report_divs:
                        txt = d.text.strip()
                        if not txt:
                            continue
                        if not edu_period_val and date_pattern.search(txt):
                            edu_period_val = txt
                        top = int(d.value_of_css_property('top').replace('px', ''))
                        left = int(d.value_of_css_property('left').replace('px', ''))
                        items.append({'txt': txt, 'top': top, 'left': left})

                    if not items:
                        break
                    items.sort(key=lambda x: (x['top'], x['left']))
                    last_page_fingerprint = "".join([x['txt'] for x in items[:3]])
                    print(f"    - {course_main_text[:10]}... {page_count}페이지 파싱")

                    # 좌표(top)로 행 묶기
                    rows = []
                    current_row = [items[0]]
                    for idx in range(1, len(items)):
                        if abs(items[idx]['top'] - current_row[-1]['top']) < 12:
                            current_row.append(items[idx])
                        else:
                            rows.append(current_row)
                            current_row = [items[idx]]
                    rows.append(current_row)

                    page_rows = 0
                    for r in rows:
                        r.sort(key=lambda x: x['left'])
                        fixed_row = [""] * 19
                        fixed_row[16] = course_main_text
                        fixed_row[17] = edu_period_val
                        fixed_row[18] = f"{year}-{month:02d}"

                        for item in r:
                            x, val = item['left'], item['txt']
                            if x < 50:
                                fixed_row[0] = val
                            elif 50 <= x < 150:
                                if not fixed_row[1]: fixed_row[1] = val
                                else: fixed_row[2] = val
                            elif 150 <= x < 250:
                                if not fixed_row[3]: fixed_row[3] = val
                                else: fixed_row[4] = val
                            elif 250 <= x < 350:
                                fixed_row[5] = val
                            elif 350 <= x < 500:
                                if not fixed_row[6]: fixed_row[6] = val
                                else: fixed_row[7] = val
                            elif 500 <= x < 650:
                                if not fixed_row[8]: fixed_row[8] = val
                                else: fixed_row[9] = val
                            elif 650 <= x < 800:
                                if not fixed_row[10]: fixed_row[10] = val
                                else: fixed_row[11] = val
                            elif 800 <= x < 950:
                                if not fixed_row[12]: fixed_row[12] = val
                                else: fixed_row[13] = val
                            elif 950 <= x < 1050:
                                fixed_row[14] = val
                            elif x >= 1050:
                                fixed_row[15] = val

                        if fixed_row[0]:
                            new_data.append(fixed_row)
                            page_rows += 1

                    # 💡 [핵심] 병목 제거: 한 명마다 저장하던 것을 한 페이지마다 저장으로
                    if page_rows:
                        print(f"    👉 {page_count}페이지 {page_rows}명 수집")
                        save_to_excel()

                    # 다음 페이지
                    try:
                        next_button = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.ID, "next")))
                        if "m-disable" in (next_button.get_attribute("class") or ""):
                            print("    📢 마지막 페이지(버튼 비활성)")
                            break
                        driver.execute_script("arguments[0].click();", next_button)
                        page_count += 1
                        # sleep 제거: 위쪽 fingerprint 폴링이 전환을 알아서 감지함
                    except Exception:
                        print("    📢 단일 페이지 과정")
                        break

                driver.close()
                driver.switch_to.window(main_window)
                driver.switch_to.frame("kas01_005_t")

            except Exception as course_err:
                print(f"❌ 과정 수집 에러(스킵): {course_err}")
                if main_window and len(driver.window_handles) > 1:
                    driver.close()
                    driver.switch_to.window(main_window)
                    driver.switch_to.frame("kas01_005_t")
                continue

    save_to_excel()  # 최종 저장
    print("\n✨ 전체 완료")
    subprocess.run(["open", DESKTOP])

except Exception as e:
    print(f"🔥 치명적 에러: {e}")
    save_to_excel()  # 중단되더라도 지금까지 수집분 저장
finally:
    driver.quit()


# ─────────────────────────────────────────────────────────
# 📌 메뉴 자동 이동 동작 방식
# ─────────────────────────────────────────────────────────
# MENU_PATH = ["일반행정", "부서업무", "안전교육원", "수강생관리"] 의 글자를 순서대로
# 클릭합니다. 태그 종류(a/span/div/li 등)나 프레임 위치를 몰라도 글자로 찾아 누릅니다.
# - 특정 항목을 자동으로 못 찾으면 그 항목만 직접 클릭 후 엔터를 치면 이어서 진행됩니다.
# - 메뉴 글자가 위 표기와 다르면(띄어쓰기 등) MENU_PATH 의 글자를 실제와 똑같이 맞추세요.
