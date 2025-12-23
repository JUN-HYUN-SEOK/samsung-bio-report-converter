import time
import gspread
import json
import os
import pyperclip
import pyautogui
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import UnexpectedAlertPresentException, NoSuchElementException, TimeoutException


# [설정] 구글 시트 및 유니패스 정보
# ---------------------------------------------------------
JSON_FILE = 'service_account.json'
SHEET_KEY = '14lU9FloAZVmaBt4tl8-kZ2iuoM4I0BPOGVW2F0xX3Vw'
WORKSHEET_NAME = 'Scraping'
UNIPASS_URL = "https://unipass.customs.go.kr/csp/index.do"
CONFIG_FILE = 'config.json'
# ---------------------------------------------------------

def load_config(config_file=CONFIG_FILE):
    """설정 파일 로드"""
    if os.path.exists(config_file):
        with open(config_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None

def auto_login(driver, username, password, cert_password, log_func=None):
    """
    유니패스 자동 로그인 (pyautogui 기반 완전 자동화)
    
    Args:
        driver: Selenium WebDriver
        username: 유니패스 아이디
        password: 유니패스 비밀번호
        cert_password: 인증서 비밀번호
        log_func: 로그 출력 함수
    """
    def log(msg, level="INFO"):
        if log_func:
            log_func(msg, level)
        else:
            print(msg)
    
    try:
        # 1. 로그인 버튼 클릭
        log("▶ 로그인 버튼 클릭 중...")
        login_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "QUICK_MENU_btnLogin"))
        )
        login_btn.click()
        log("  ✓ 로그인 버튼 클릭 완료")
        time.sleep(2)
        
        # 2. 아이디 입력
        log("▶ 아이디 입력 중...")
        id_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "loginId"))
        )
        id_field.click()
        pyperclip.copy(username)
        id_field.send_keys(Keys.CONTROL, 'v')
        log(f"  ✓ 아이디 입력 완료: {username}")
        time.sleep(1)
        
        # 3. 비밀번호 입력
        log("▶ 비밀번호 입력 중...")
        pw_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "loginPwd"))
        )
        pw_field.click()
        pyperclip.copy(password)
        pw_field.send_keys(Keys.CONTROL, 'v')
        log("  ✓ 비밀번호 입력 완료")
        time.sleep(1)
        
        # 4. 자동 로그인 체크박스
        log("▶ 자동 로그인 체크박스 체크 중...")
        try:
            auto_login_checkbox = driver.find_element(By.ID, "BrwsAtdcLgnYn")
            if not auto_login_checkbox.is_selected():
                auto_login_checkbox.click()
                log("  ✓ 자동 로그인 체크박스 체크 완료")
        except:
            log("  ⚠️ 체크박스를 찾을 수 없습니다", "WARNING")
        time.sleep(1)
        
        # 5. 공인인증서 로그인 버튼 클릭
        log("▶ 공인인증서 로그인 버튼 클릭 중...")
        try:
            cert_login_btn = driver.find_element(By.XPATH, "//button[contains(text(), '공인인증서') or contains(text(), '인증서')]")
            cert_login_btn.click()
        except:
            login_submit = driver.find_element(By.XPATH, "//button[@type='submit' or contains(@class, 'login')]")
            login_submit.click()
        log("  ✓ 공인인증서 로그인 버튼 클릭 완료")
        time.sleep(3)
        
        # 6. 권한 팝업 처리 (pyautogui)
        log("▶ 권한 팝업 처리 중...")
        time.sleep(3)
        try:
            log("  ✓ TAB+TAB+ENTER 키 입력 시작...")
            pyautogui.press('tab')
            time.sleep(0.3)
            pyautogui.press('tab')
            time.sleep(0.3)
            pyautogui.press('enter')
            log("  ✓ 권한 팝업 처리 완료 (TAB+TAB+ENTER)")
        except Exception as e:
            log(f"  ⚠️ 팝업 처리 중 오류: {e}", "WARNING")
        time.sleep(2)
        
        # 7. 인증서 창 전환 대기
        log("▶ 공인인증서 선택 창 대기 중...")
        time.sleep(5)
        
        # 8. 하드디스크/이동식 버튼 클릭 (pyautogui)
        log("▶ 하드디스크/이동식 버튼 클릭 중...")
        time.sleep(2)
        pyautogui.press('tab')
        time.sleep(0.2)
        pyautogui.press('tab')
        time.sleep(0.2)
        pyautogui.press('enter')
        log("  ✓ 하드디스크/이동식 버튼 클릭 완료")
        time.sleep(2)
        
        # 9. MagicLine4NX 팝업 처리
        log("  ▶ MagicLine4NX 팝업 처리 중...")
        time.sleep(2)
        pyautogui.press('tab')
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(0.3)
        pyautogui.press('tab')
        time.sleep(0.3)
        pyautogui.press('enter')
        log("  ✓ MagicLine4NX 팝업 처리 완료")
        time.sleep(3)
        
        # 10. 하드디스크 재선택
        log("  ▶ 하드디스크 재선택 중...")
        time.sleep(2)
        pyautogui.press('tab')
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(0.3)
        pyautogui.press('enter')
        log("  ✓ 하드디스크 재선택 완료")
        time.sleep(2)
        
        # 11. 인증서 선택 (TAB 4번 + DOWN 1번)
        log("  ▶ 우신관세사법인 인증서 선택 중...")
        time.sleep(2)
        for i in range(4):
            pyautogui.press('tab')
            time.sleep(0.1)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('enter')
        log("  ✓ 우신관세사법인 인증서 선택 완료")
        time.sleep(2)
        
        # 12. 인증서 비밀번호 입력
        log("  ▶ 인증서 비밀번호 입력 중...")
        time.sleep(2)
        pyautogui.press('tab')
        time.sleep(0.2)
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.write(cert_password, interval=0.1)
        log("  ✓ 비밀번호 입력 완료")
        time.sleep(0.5)
        
        # 13. ENTER로 로그인 완료
        log("  ▶ ENTER로 로그인 완료...")
        pyautogui.press('enter')
        log("  ✓ 로그인 완료!")
        time.sleep(2)
        
        # 14. 비밀번호 저장 팝업 닫기
        log("  ▶ 비밀번호 저장 팝업 닫는 중...")
        pyautogui.press('esc')
        time.sleep(0.5)
        pyautogui.press('esc')
        log("  ✓ 팝업 닫기 완료")
        time.sleep(2)
        
        log("  ✅ 로그인 성공!", "SUCCESS")
        return True
        
    except Exception as e:
        log(f"❌ 로그인 중 오류 발생: {e}", "ERROR")
        return False

def connect_google_sheet(json_file=None, sheet_key=None, worksheet_name=None):
    """구글 시트 연결"""
    if json_file is None:
        json_file = JSON_FILE
    if sheet_key is None:
        sheet_key = SHEET_KEY
    if worksheet_name is None:
        worksheet_name = WORKSHEET_NAME
    
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_file, scope)
    client = gspread.authorize(creds)
    sh = client.open_by_key(sheet_key)
    worksheet = sh.worksheet(worksheet_name)
    return worksheet

def run_automation_with_callback(config=None, callback=None):
    """
    콜백 함수를 받아 GUI와 통신하는 자동화 함수
    
    Args:
        config: 설정 딕셔너리 (None이면 기본값 사용)
        callback: 이벤트 콜백 함수 callback(event_type, data)
    """
    def log(message, level="INFO"):
        """로그 출력 (콘솔 및 콜백)"""
        print(message)
        if callback:
            callback("log", {"message": message, "level": level})
    
    def update_progress(current, total):
        """진행률 업데이트"""
        if callback:
            callback("progress", {"current": current, "total": total})
    
    def update_status(message):
        """상태 메시지 업데이트"""
        if callback:
            callback("status", {"message": message})
    
    # 설정 사용 (config가 없으면 기본값)
    if config is None:
        config = {
            'json_file': JSON_FILE,
            'sheet_key': SHEET_KEY,
            'worksheet_name': WORKSHEET_NAME,
            'unipass_url': UNIPASS_URL
        }
    
    # 1. 구글 시트 연결 및 데이터 읽기
    log("▶ 구글 시트에 연결 중...")
    ws = connect_google_sheet(config.get('json_file', JSON_FILE), 
                              config.get('sheet_key', SHEET_KEY),
                              config.get('worksheet_name', WORKSHEET_NAME))
    all_records = ws.get_all_values()
    
    target_rows = []
    for idx, row in enumerate(all_records[1:], start=2):
        if len(row) < 2:
            continue
        
        # B열 값을 코드로 사용 (전체 값 사용)
        code_value = row[1].strip()
        
        # B열이 비어있지 않으면 조회 대상에 추가
        if code_value:
            # A열에 업체명이 있으면 사용, 없으면 B열 값을 표시용으로 사용
            company_name = row[0].strip() if row[0].strip() else code_value
            target_rows.append({'row_idx': idx, 'name': company_name, 'code': code_value})
            log(f"  ✓ 조회 대상: [{company_name}] 코드: {code_value}")
    
    if not target_rows:
        log("조회할 업체 코드가 시트에 없습니다.", "WARNING")
        return

    log(f"\n총 {len(target_rows)}개 업체를 조회합니다.\n", "INFO")
    update_progress(0, len(target_rows))

    # 2. 브라우저 실행
    log("▶ Chrome 브라우저를 시작합니다...")
    update_status("브라우저 시작 중...")
    
    # Chrome 옵션 설정
    chrome_options = Options()
    
    # 권한 요청 팝업 자동 거부 설정
    prefs = {
        "profile.default_content_setting_values.media_stream": 2,  # 카메라/마이크 거부
        "profile.default_content_setting_values.geolocation": 2,   # 위치 거부
        "profile.default_content_setting_values.notifications": 2,  # 알림 거부
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    # 팝업 차단 옵션 추가
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-features=BlockInsecurePrivateNetworkRequests")
    chrome_options.add_argument("--disable-features=PrivateNetworkAccessSendPreflights")
    
    # 자동화 감지 회피
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # 설정 파일에서 Chrome 프로필 경로 가져오기
    app_config = load_config()
    if app_config and app_config.get('unipass', {}).get('use_chrome_profile') and app_config.get('unipass', {}).get('chrome_profile_path'):
        profile_path = app_config['unipass']['chrome_profile_path']
        if os.path.exists(os.path.dirname(profile_path)):
            chrome_options.add_argument(f"user-data-dir={profile_path}")
            log(f"  ✓ Chrome 프로필 사용: {profile_path}")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    try:
        driver.get(config.get('unipass_url', UNIPASS_URL))
        log("\n▶ 유니패스 페이지 로딩 완료")
        time.sleep(2)
        
        # 자동 로그인 시도
        if app_config and app_config.get('unipass', {}).get('username'):
            username = app_config['unipass']['username']
            password = app_config['unipass'].get('password', '')
            cert_password = app_config['unipass'].get('cert_password', '')

            
            if username and password:
                log("\n▶ 자동 로그인을 시작합니다...")
                update_status("자동 로그인 중...")
                auto_login(driver, username, password, cert_password, log)
            else:
                log("\n▶ [설정 필요] config.json에 username과 password를 입력해주세요.", "WARNING")
                log("▶ [로그인 대기] 수동으로 로그인 후 콘솔에서 [Enter]를 누르세요.", "WARNING")
                update_status("로그인 대기 중... (콘솔에서 Enter를 누르세요)")
                input()
        else:
            log("\n▶ [로그인 대기] 수동으로 로그인 후 콘솔에서 [Enter]를 누르세요.", "WARNING")
            update_status("로그인 대기 중... (콘솔에서 Enter를 누르세요)")
            input() 

        # 3. 시트의 업체 리스트 순회하며 조회
        for idx, item in enumerate(target_rows, 1):
            try:
                log(f"\n{'='*50}")
                log(f"[{idx}/{len(target_rows)}] {item['name']} 처리 시작")
                log(f"{'='*50}")
                update_status(f"{item['name']} 조회 중... ({idx}/{len(target_rows)})")
                ws.update_cell(item['row_idx'], 7, "조회중...") 
                
                # [핵심 수정] 매번 메뉴를 새로 클릭하여 세션 토큰을 갱신함
                # 메뉴 검색창 찾기
                try:
                    search_box = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "menuSearch")))
                    search_box.clear()
                    search_box.send_keys("월별 납부한도액 사용내역")
                    search_box.send_keys(Keys.RETURN)
                    
                    time.sleep(2)
                    
                    # 메뉴 링크 클릭
                    try:
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, "월별 납부한도액 사용내역"))).click()
                    except:
                        driver.find_element(By.PARTIAL_LINK_TEXT, "월별 납부한도액").click()
                        
                except UnexpectedAlertPresentException:
                    try:
                        driver.switch_to.alert.accept()
                    except:
                        pass
                except Exception as e:
                    log(f"   -> 메뉴 이동 중 오류: {e}", "WARNING")

                # 팝업 처리 (혹시 뜰 경우)
                time.sleep(2)
                try:
                    driver.switch_to.alert.accept()
                except:
                    pass

                # 화면 로딩 대기
                WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "MYC0119014Q_ecm")))

                # 업체 선택
                try:
                    select_element = Select(driver.find_element(By.ID, "MYC0119014Q_ecm"))
                    select_element.select_by_value(item['code'])
                except NoSuchElementException:
                    log(f"   -> [경고] 업체 코드({item['code']})를 찾을 수 없습니다.", "WARNING")
                    ws.update_cell(item['row_idx'], 7, "코드없음")
                    continue

                # 기존 결과 필드 비우기
                try:
                    mg_amt_elem = driver.find_element(By.ID, "mgAmt")
                    apnt_xpir_dt_elem = driver.find_element(By.ID, "apntXpirDt")
                    driver.execute_script("arguments[0].textContent = '';", mg_amt_elem)
                    driver.execute_script("arguments[0].textContent = '';", apnt_xpir_dt_elem)
                except:
                    pass

                # 조회 버튼 클릭
                try:
                    search_btn = driver.find_element(By.XPATH, "//footer//button[contains(., '조회')]")
                    driver.execute_script("arguments[0].click();", search_btn)
                except:
                    driver.switch_to.active_element.send_keys(Keys.ENTER)
                
                log(f"   ⏳ 데이터 로딩 대기 중... (최대 30초)")
                update_status(f"{item['name']} 데이터 로딩 중...")
                
                # 데이터 로딩 대기
                try:
                    WebDriverWait(driver, 30).until(
                        lambda d: d.find_element(By.ID, "mgAmt").text.strip() != "" or d.find_element(By.ID, "mgAmt").get_attribute("value")
                    )
                except TimeoutException:
                    log("   -> 데이터 로딩 시간 초과 (또는 데이터 없음)", "WARNING")

                # 데이터 추출 함수
                def get_text_or_value(element_id):
                    try:
                        el = driver.find_element(By.ID, element_id)
                        txt = el.text
                        if txt:
                            txt = txt.strip()
                        
                        if not txt:
                            val = el.get_attribute("value")
                            if val:
                                txt = val.strip()
                            else:
                                txt = ""
                        return txt
                    except Exception as e:
                        return f"[못찾음]"

                limit_amt = get_text_or_value("mgAmt")
                use_amt = get_text_or_value("mgUseAmt")
                balance_amt = get_text_or_value("mgBlAmt")
                expire_date = get_text_or_value("apntXpirDt")
                
                log(f"   ✓ 추출 완료: 한도={limit_amt}, 사용={use_amt}, 잔액={balance_amt}, 만료={expire_date}")
                update_status(f"{item['name']} 시트 업데이트 중...")

                cell_list = ws.range(f'C{item["row_idx"]}:F{item["row_idx"]}')
                cell_list[0].value = limit_amt
                cell_list[1].value = use_amt
                cell_list[2].value = balance_amt
                cell_list[3].value = expire_date
                ws.update_cells(cell_list)
                
                ws.update_cell(item['row_idx'], 7, "완료")
                
                current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                ws.update_cell(item['row_idx'], 8, current_time)
                
                log(f"   ✅ [{item['name']}] 완료! (시간: {current_time})", "SUCCESS")
                update_progress(idx, len(target_rows))

            except Exception as e:
                log(f"❌ Error: {e}", "ERROR")
                ws.update_cell(item['row_idx'], 7, f"에러: {str(e)}")
                update_progress(idx, len(target_rows))

        log("\n" + "="*60, "SUCCESS")
        log("▶ 모든 작업이 완료되었습니다!", "SUCCESS")
        log("="*60, "SUCCESS")

    except Exception as e:
        log(f"시스템 오류: {e}", "ERROR")
    finally:
        driver.quit()


def run_automation():
    """기존 호환성을 위한 래퍼 함수"""
    run_automation_with_callback()


if __name__ == "__main__":
    run_automation()
    
