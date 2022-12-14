from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
from datetime import datetime
from selenium.webdriver.support.ui import Select
import re
import requests
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from twocaptcha import TwoCaptcha
import os
import glob
import numpy as np
from urllib.parse import urljoin

class jobkrSCR:
    def __init__(self):
        self.api_key = '5ecc9cfa3451a7e4d3a000fcac87d3d1'
        # 메뉴 선택
        self.input_num = int()
        # url 초기화
        self.init_url = "https://www.jobkorea.co.kr/recruit/joblist?menucode=local&localorder=1"

        self.driver = webdriver.Chrome(executable_path= './chromedriver.exe')
        self.driver.implicitly_wait(10)
        self.driver.set_page_load_timeout(10)
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36'}

        # ------------------- 1차 수집시 필요 생성자 ------------------------
        # 1차 결과 저장
        self.temp_1st = pd.DataFrame(columns=["co_name", "co_link"])
        self.results_1st = pd.DataFrame(columns=["co_name", "co_link"])

        # 1차 결과 검색시 페이지 이동을 위한 현재 페이지 번호
        self.cur_page = 1
        self.dup_idx = -1

        # 공고 등록일 정규식
        self.p = re.compile('^.*.일 전 등록$')

        # ------------------- 2차 수집시 필요 생성자 ------------------------
        self.search_df_2nd = ''
        self.co_items = []
        self.none_data = ['None', 'None', 'None', 'None','None']
        # -------------------------------------------------------------------
        self.sch = ''
        self.pattern = re.compile(r'([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)')

    # 메뉴 선택
    def print_menu(self):
        print('-------------------------------------')
        print('1. 1차 수집 시작')
        print('2. 2차 수집 시작')
        print('3. 이메일 스크랩 시작')
        print('4. 3번 결과 통합 + 데이터 전처리')
        print('5. 종료')
        print('-------------------------------------')

        self.input_num = input_num = int(input('메뉴 번호를 입력하세요: '))

        if input_num == 1:
            print('1차 수집 시작')
            self.run_1st_script()
        elif input_num == 2:
            print('2차 수집 시작')
            self.run_2nd_script()
        elif input_num == 3:
            print('이메일 스크랩 시작')
            self.run_3rd_script()
        elif input_num == 4:
            print('모든 데이터 통합 + 데이터 전처리 시작')
            self.run_4th_script()
        elif input_num == 5:
            return False
        else:
            print('번호를 제대로 입력해주세요')
        return True

    def get_link(self):
        try:
            self.driver.get(self.init_url)
        except:
            self.driver = webdriver.Chrome('./chromedriver.exe')
            time.sleep(2)
            self.driver.get(self.init_url)

    # ----------------- 1차 수집 관련 함수 --------------------------
    # 공고 필터링 & 1차수집 시작
    def run_1st_script(self):
        self.get_link()
        self.search_filter()
        self.scrap_1st()

    # 공고 조회시 중소기업 벤처기업 필터링 및 등록순으로 정렬
    def search_filter(self):
        # 검색조건 추가
        scripts = """
        document.querySelector("#toolitems").innerHTML=`<li class='cotype item' data-value='15' data-group='cotype15' data-type=''>
            <button type='button'>중소기업<span class='ico'>삭제</span></button>
        </li>
        <li class="cotype item" data-value="7" data-group="cotype7" data-type="">
            <button type="button">벤처기업<span class="ico">삭제</span></button>
        </li>
        `;
        document.querySelector(".resultSet").style.display = 'block';
        """
        self.driver.execute_script(scripts)

        # 조회하기 버튼 클릭
        co_type_btn = self.driver.find_element(by=By.CSS_SELECTOR, value="#dev-btn-search")
        co_type_btn.send_keys(Keys.ENTER)

        # 공고 조회 대기
        time.sleep(5)

        # 최신 등록순으로 정렬 (위험 있는 코드, 셀렉트 박스를 찾지못함)
        select = Select(self.driver.find_element(by=By.ID, value='orderTab'))
        select.select_by_value('2')

        # 정렬 후 대기
        time.sleep(2)

    # 1차 수집(회사명/ 회사 잡코리아링크/ 업로드 시간 수집)
    def scrap_1st(self):
        while True:
            start = time.time()
            # 각 회사별 잡코리아링크/회사명 수집
            tr = self.driver.find_elements(by=By.CSS_SELECTOR, value="#dev-gi-list>div>.tplList>table>tbody>tr>td>a")
            jk_co_link = [i.get_attribute("href") for i in tr]
            jk_co_name = [i.text for i in tr]

            # 공고 업로드 시간 조회
            time_raw = self.driver.find_elements(by=By.CSS_SELECTOR, value="#dev-gi-list>div>.tplList>table >tbody>tr>td.odd>span.time.dotum")
            time_arr = [t.text for t in time_raw]

            # 공고 업로드 시간 검증(N일 전 업로드 제외)
            try:
                self.dup_idx = [bool(self.p.match(t)) for t in time_arr].index(True)
            except ValueError as m:
                # N시간 전 인 경우 패스
                pass
            else:
                # 공고 목록 중 N일 전인 항목이 있을 경우 슬라이싱
                if (self.dup_idx > 0):
                    jk_co_name = jk_co_name[:self.dup_idx]
                    jk_co_link = jk_co_link[:self.dup_idx]
                    #time_arr = time_arr[:self.dup_idx]
            finally:
                # 업로드 시간에 따라 필터링 완료 후 결과 리스트에 추가
                temp_dict = {
                    "co_name": jk_co_name,
                    "co_link": jk_co_link
                }
                self.temp_1st = pd.DataFrame(temp_dict)
                self.results_1st = pd.concat([self.results_1st, self.temp_1st])
                # 중복 제거
                self.results_1st = self.results_1st.drop_duplicates(['co_name', 'co_link'])
                # 완전히 수집 후 종료
                if (self.dup_idx >= 0):
                    self.results_1st.to_excel("./rawdata/기업정보_1차_수집" + datetime.today().strftime("%Y%m%d%H%M%S") + ".xlsx", index=False)
                    break

            # 끝나고 다음페이지
            self.cur_page = self.cur_page + 1

            # 페이지 이동 스크립트
            if (self.cur_page % 10 == 1):
                # 다음페이지 버튼 클릭
                self.driver.find_element(by=By.CLASS_NAME, value="btnPgnNext").send_keys(Keys.ENTER)
                self.cur_page = self.cur_page % 10
            else:
                pagenation_link = self.driver.find_element(by=By.XPATH,value="//*[@id='dvGIPaging']/div/ul/li[%d]/a" % self.cur_page)
                pagenation_link.send_keys(Keys.ENTER)
            # 목록 당 소요시간 출력
            print("한 페이지 수집 완료 | time taken : ", int(time.time() - start), "s")

    # ----------------- 2차 수집 관련 함수 --------------------------
    # 2차 수집 세팅 & 스크립트 시작
    def run_2nd_script(self):
        self.set_corp()
        self.scrap_2nd()

    # 2차 수집에 필요한 세터
    def set_corp(self):
        path = self.load_dataset(1)
        try:
            self.search_df_2nd = pd.read_excel(path)
        except:
            print('파일을 찾을 수 없습니다. 1차 수집을 먼저 진행 해주세요')

    # 2차 수집 상세 스크립트
    def scrap_2nd(self):
        for href in self.search_df_2nd['co_link']:
            start = time.time()
            co_info_list = []
            # 기업별 잡코리아 링크 응답한 경우 vs 응답하지 못한 경우
            try:
                self.driver.get(href)
            except:
                co_info_list = self.none_data
                self.co_items.append(co_info_list)
                pass

            # 기업별 잡코리아 링크 응답해도 기업정보 데이터를 찾는 경우 vs 못 찾는 경우
            try:
                # 데이터 테이블 추출
                td = self.driver.find_elements(by=By.CSS_SELECTOR,
                                               value='.table-basic-infomation-primary .field .value')
                td_text = [i.text for i in td]

                # 기업정보 목록 산업분류/사원수/기업구분/홈페이지/주소
                co_info_list = [td_text[0], td_text[1], td_text[2], td_text[-2], td_text[-1]]

                if (not co_info_list):
                    co_info_list = self.none_data

                self.co_items.append(co_info_list)
            except:
                try:
                    # 캡차 있는지 체크
                    self.driver.execute_script('showCaptchaImage(true)')

                    # 캡차 있다면 API 실행
                    self.captcha_next_btn()
                    self.save_img()
                    self.solve_captcha()

                    # 성공적으로 캡차를 뚫은 경우 데이터 수집
                    td = self.driver.find_elements(by=By.CSS_SELECTOR,
                                                   value='.table-basic-infomation-primary .field .value')
                    td_text = [i.text for i in td]

                    co_info_list = [td_text[0], td_text[1], td_text[2], td_text[-2], td_text[-1]]
                except:
                    # 캡차도 없는데 기업정보도 없는 경우 패스
                    pass
                # 위와 같은 작업 후 정보가 없는 경우 none 처리
                if (not co_info_list):
                    co_info_list = self.none_data
                # 리스트에 담기
                self.co_items.append(co_info_list)

                pass

            print("time taken : ", int(time.time() - start), " site: ", self.co_items[-1][-2])
        co_items_df = pd.DataFrame(self.co_items)
        snd_results = pd.concat([self.search_df_2nd, co_items_df], axis=1)
        snd_results.to_excel("./rawdata/기업정보_2차_수집" + datetime.today().strftime("%Y%m%d%H%M%S") + ".xlsx", index=False)

    def solve_captcha(self):
        solver = TwoCaptcha(self.api_key)
        try:
            result = solver.normal('./captcha.png')
            result_code = result['code']

            print(result_code)
            input_area = self.driver.find_element(by=By.ID, value="txtInputText")
            input_area.send_keys(result_code)

            submit_btn = self.driver.find_element(by=By.ID, value="btnInput")
            submit_btn.send_keys(Keys.ENTER)
        except:
            pass

    def save_img(self):
        time.sleep(3)
        img = self.driver.find_element(by=By.ID, value='imgCaptcha').screenshot_as_png
        with open('./captcha.png', 'wb') as file:
            file.write(img)

    def captcha_next_btn(self):
        self.driver.execute_script('showCaptchaImage(true)')

    # ----------------- 최종 수집(이메일 스크랩) --------------------------
    # 3차 수집 시작
    def run_3rd_script(self):
        self.set_corp_for_3rd()
        self.scrap_3rd()

    def set_corp_for_3rd(self):
        path = self.load_dataset(2)
        try:
            self.sch = pd.read_excel(path)
        except:
            print('파일을 찾을 수 없습니다. 2차 수집을 먼저 진행 해주세요')

    # 3차 수집 상세 스크립트
    def scrap_3rd(self):
        comp_email = []

        for u in self.sch[3]:
            start = time.time()
            email = 'None'
            if u!='-':
                try:
                    response = self.connect_url(u, 10) #기업링크에 연결 요청, 10초가 지나면 연결하지않음
                    soup = BeautifulSoup(response.content, "html.parser")
                except:
                    email = 'None' #기업링크 연결 요청이 안 될 경우 이메일 수집 포기
                else:
                    email = self.soup_scrap(bs = soup, url = u)
                    if(email=='None' or email=='' or email==False):
                        redirect_url = self.redirect_url_return(u)
                        if(redirect_url):
                            try:
                                response = self.connect_url(redirect_url, 10)  # 기업링크에 연결 요청, 10초가 지나면 연결하지않음
                                soup = BeautifulSoup(response.content, "html.parser")
                            except:
                                email = 'None'
                            else:
                                email = self.soup_scrap(bs = soup, url = redirect_url)
                        else:
                            email = 'None'
                finally:
                    comp_email.append(email)
            else:
                comp_email.append(email)
            print(u, " : ", int(time.time() - start), "s, email: ", email)

        # 이메일 수집 결과와 2차 수집 데이터프레임 병합
        last_em_df = pd.DataFrame(comp_email)

        # 폴더 생성
        dir_path = datetime.today().strftime("%Y-%m-%d_이메일_수집_결과")
        self.createDirectory(dir_path)

        # 이메일 수집 결과 전처리&저장
        self.sch = pd.concat([self.sch, last_em_df], axis=1)

        self.sch = self.pretreatment(self.sch)
        self.sch.to_excel("./results/" + dir_path + "/이메일_수집_결과" + datetime.today().strftime("%Y%m%d%H%M%S") + ".xlsx",
                          index=False)

    #최종 산출물 전처리 데이터 병합&중복처리
    def run_4th_script(self):
        # 초기화
        merge_df = pd.DataFrame()

        # 파일명이 기업정보_최종_산출물로 시작하는 엑셀 파일(.xlsx) 전부 로드해오기
        file_path = "./results"
        file_list = [f"./results/{file}/" for file in os.listdir(file_path)
                     if re.match("\d{4}-\d{2}-\d{2}_이메일_수집_결과", file)]

        # results 폴더에 있는 엑셀 파일 병합
        for file_name in file_list:
            name = file_name + "*.xlsx"
            file_df = pd.read_excel(os.path.normpath(glob.glob(name)[0]))
            merge_df = merge_df.append(file_df, ignore_index=True)
            merge_df = merge_df.drop_duplicates(['회사명', '이메일'])

        merge_df = self.pretreatment(merge_df)

        # 최종 저장
        merge_df.to_excel("./전처리 완료 결과/파일 통합 + 이메일 전처리" + datetime.today().strftime("%Y%m%d%H%M%S") + ".xlsx", index=False)


    # 최근 파일 불러오기
    def load_dataset(self, idx):
        if not idx:
            return -1

        raw_path = "rawdata/"
        files = glob.glob(raw_path+'*.xlsx')
        files = [os.path.basename(i) for i in files]

        co_file = list(filter(lambda x: re.match('기업정보_{0}차_수집.*\.xlsx'.format(idx), x), files))
        files_Path = raw_path  # 파일들이 들어있는 폴더
        file_name_and_time_lst = []

        # 해당 경로에 있는 파일들의 생성시간을 함께 리스트로 넣어줌.
        for f_name in co_file:
            written_time = os.path.getctime(f"{files_Path}{f_name}")
            file_name_and_time_lst.append((f_name, written_time))

        # 생성시간 역순으로 정렬하고,
        sorted_file_lst = sorted(file_name_and_time_lst, key=lambda x: x[1], reverse=True)

        try:
            # 가장 앞에 있는 파일을 넣어준다.
            recent_file = sorted_file_lst[0]
            recent_file_name = recent_file[0]
        except:
            return -2

        return raw_path + recent_file_name

    # 메타태그로 인해 리다이렉트시 새 url 리턴
    def redirect_url_return(self, url):
        # 기업링크에 연결 요청, 10초가 지나면 연결하지않음
        response = self.connect_url(url, 10)
        if(not response):
            return False
        soup = BeautifulSoup(response.content, "html.parser")

        try:
            result = soup.find_all("meta", attrs={"http-equiv": "Refresh"})[0]
        except:
            result = ''

        if result:
            wait, text = result["content"].split(";")
            if text.strip().lower().startswith("url="):
                sub_url = re.sub('url=', "", text, flags=re.IGNORECASE)
                if (url[-1] == '/'):
                    return_url = url.join(sub_url)
                else:
                    return_url = '/'.join([url, sub_url])
                return_url = return_url.replace(" ", "")

                return return_url
            else:
                return False
        else:
            return False

    # 데이터 크롤링
    def soup_scrap(self, bs, url, res=''):
        email = ''
        # 불필요한 태그 제거
        script_tag = bs.find_all(['script', 'style', 'head'])

        for script in script_tag:
            script.extract()

        # html 문서 행 분리
        content = bs.get_text('\n', strip=True)
        html_split = content.split("\n")

        # 메일 추출
        email = self.find_email(html_split)

        # 문서 행 분리 시 컨텐츠가 비어있을경우
        if(html_split=='' or email==''):
            try:
                self.driver.get(url)
                html = self.driver.page_source
                bs = BeautifulSoup(html, 'html.parser')
            except:
                bs = self.change_parser(res)
            else:
                if (len(self.driver.window_handles) > 1):
                    self.close_popup(self.driver)
            finally:
                content = bs.get_text('\n', strip=True)
                html_split = content.split("\n")

                # 메일 추출
                email = self.find_email(html_split)

        # iframe/frame 체크
        if (email=='' and bs.find(['frame','iframe'],src=True)):
            frame = bs.find(['frame','iframe'], src=True)
            try:
                new_url = urljoin(url, frame['src'])

                bs = BeautifulSoup(self.connect_url(new_url,10).content, "html.parser")

                content = bs.get_text('\n', strip=True)
                html_split = content.split("\n")

                # 메일 추출
                email = self.find_email(html_split)
            except:
                pass

        # html코드를 스플릿 하지 못해서 못 찾은 경우
        if (email == ''):
            email = self.find_email(bs, 'prettify')

        # 그래도 없을 경우 a태그에서 추출
        if (email==''):
            try:
                attrs = [i.attrs['href'] for i in bs.find_all("a", href=True)]
                email = list(filter(lambda x: re.findall(self.pattern, x), attrs))[0].replace("mailto:", "")
            except:
                pass

        return email


    # html 문서 변환 형식 변경
    def change_parser(self,res):
        soup = BeautifulSoup(res, "html5lib")
        return soup


    # 웹 연결 요청
    def connect_url(self, url, timeout):
        try:
            response = requests.get(url, timeout=timeout, headers=self.headers, verify=False)
        except:
            response = False
        return response


    # 스플릿된 뷰티풀 수프 객체를 나누고, 이메일 정규식을 적용하여 이메일 주소를 리턴
    def find_email(self, html_split, opt=''):
        email = ''
        if(opt=='prettify'):
            new_split = html_split.prettify().split("\n")
        else:
            new_split = html_split

        for row in new_split:
            mail = re.findall(self.pattern, row)
            if (mail):
                email = mail[0]
                return email
        return email

    # 팝업 닫기
    def close_popup(self, driver):
        tabs = driver.window_handles
        last_tab = driver.window_handles[-1]
        try:
            while len(tabs)!=1:
                driver.switch_to.window(last_tab)
                driver.close()
                tabs = self.driver.window_handles
            driver.switch_to.window(tabs[0])
        except:
            pass
        finally:
            driver.switch_to.window(tabs[0])

    # 디렉토리 생성
    def createDirectory(self, directory):
        try:
            if not os.path.exists("./results/"+directory):
                os.makedirs("./results/"+directory)
        except OSError:
            print("Error: Failed to create the directory.")

    # 데이터프레임 전처리
    def pretreatment(self, df):
        new_df = df

        # 쓰레기 텍스트 목록 설정
        junk_txt = "Email|e-mail|E-mail|email.|Email.|email:|Email:|e-mail:|E-mail:|e-mail.|E-mail.|E-MAIL|.email|.Email|EMAIL|FAX|Tel.|Tel:|tel.|TEL.|tel:|TEL:|Copyrights|copyrights|COPYRIGHTS|Copyright|COPYRIGHT|copyright|E."

        # 컬럼 이름 설정
        column_name = ['회사명', '산업분류', '사원수', '기업규모', '홈페이지', '주소', '이메일']

        if('co_link' in new_df.columns):
            new_df = new_df.drop(columns='co_link', axis=1)
        new_df.columns = column_name

        # 이메일 정규식
        new_df['이메일'] = new_df['이메일'].str.replace(junk_txt, '', regex=True)

        # 이메일이 없는 셀 제거
        new_df = new_df.loc[~new_df['이메일'].isin(['None', ' ', None, '', np.NaN])]

        # 파일 이름이(~.jpg, ~.png) 검출된 경우 제거 & font 파일이 검출된 경우
        new_df = new_df[~new_df['이메일'].str.match('(^.*.(jpg|png)$)|(^.*.\d$)')]
        new_df = new_df[~new_df['이메일'].str.contains('font')]

        return new_df

obj = jobkrSCR()
while True:
    status = obj.print_menu()
    if (not status):
        break
