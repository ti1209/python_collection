# -*- coding: utf-8 -*-
import requests
import urllib
import xlsxwriter
import smtplib
import pandas as pd
import re
import logging
import time
import random
import json
import math

from bs4 import BeautifulSoup
from urllib import parse
from datetime import date, datetime, timedelta
from string import Template
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.utils import formatdate, formataddr
from email import encoders
from difflib import SequenceMatcher

logging.basicConfig(filename="crawl.log", format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', filemode='a', level=logging.DEBUG)

wordlist = ['자체감사', '외부감사', '회계감사', '감사시스템', '내부감사', '내부회계', '모니터링', '회계감리', '회계법인', '회계정산', '위탁정산', 'IT감사', '전산감사', 'ESG', \
    '내부통제', '감사위원', '지배구조', '위험관리', 'e감사', '품질평가', '연구개발혁신법', '회계연도', '법인세', '국세', '지방세', '세무', '전략', '감사체계', '감사품질', '직무평가', \
    '감사정보', '감사행정', '감사모니터링', '정보보호']

heute = date.today().strftime('%Y%m%d')
onemonth = (date.today() - timedelta(30)).strftime('%Y%m%d')
naver_date = (date.today() - timedelta(1)).strftime('%Y.%m.%d')
naver_another = (date.today() - timedelta(1)).strftime('%Y%m%d')
file_date = (date.today() - timedelta(1)).strftime('%Y_%m_%d')

filename = f'g2b_and_Naver_{file_date}.xlsx'
workbook = xlsxwriter.Workbook(filename)

cell_center = workbook.add_format({'align': 'center'})
cell_bold = workbook.add_format({'bold': True})
cell_url = workbook.add_format({'color': 'black', 'underline': False})

def g2b():
    worksheet = workbook.add_worksheet('나라장터 입찰공고')

    worksheet.write('A1', '실행시각', cell_bold)
    worksheet.write('B1', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

    worksheet.write('A2', '키워드', cell_bold)
    worksheet.write('B2', '업무구분', cell_bold)
    worksheet.write('C2', '사업명', cell_bold)
    worksheet.write('D2', '사업번호', cell_bold)
    worksheet.write('E2', '사업일자', cell_bold)
    worksheet.write('F2', '공고기관', cell_bold)
    worksheet.write('G2', '수요기관', cell_bold)
    worksheet.write('H2', '공고일자', cell_bold)

    worksheet.freeze_panes(2, 0)

    worksheet.set_column('A:A', 12)
    worksheet.set_column('B:B', 13)
    worksheet.set_column('C:C', 65)
    worksheet.set_column('D:D', 15)
    worksheet.set_column('E:E', 10)
    worksheet.set_column('F:F', 30)
    worksheet.set_column('G:G', 30)
    worksheet.set_column('H:H', 10)

    df = pd.DataFrame(columns = ['키워드', '업무구분', '사업명', '사업번호', '사업일자', '공고기관', '수요기관', '공고일자'])

    num = 0

    for order in wordlist:

        total = 0
        page_cycle = 1
        fin_num = 100
        
        try:
            while True:
                url = 'https://www.g2b.go.kr/fi/fiu/fiua/UntySrch/srchUntyTotal.do'

                params = {
                    "Accept" : "application/json",
                    "Accept-Encoding" : "gzip, deflate, br, zstd",
                    "Accept-Language" : "ko,en;q=0.9,en-US;q=0.8",
                    "Content-Length" : "1308",
                    "Content-Type" : "application/json;charset=UTF-8",
                    "Cookie" : "WHATAP=x6dda94o0fp8k9; XTVID=A2502171555215976; xloc=1920X1080; _harry_lang=ko; _harry_fid=hh-1741997108; JSESSIONID=MGQ5M2U4Y2UtYmQyYS00ZGI5LTk2YTUtNThlMTM5MzQ2ZWRk; Path=/; infoSysCd=%EC%A0%95010029; _harry_ref=https%3A//www.bing.com/; _harry_url=https%3A//www.g2b.go.kr/; system_language=ko; poupR23AB0000013499=done; poupR23AB0000013437=done; poupR23AB00000134104=done; lastAccess=1741660223230; globalDebug=false; XTSID=A250311135409452998; _harry_hsid=A250311135646686997; _harry_dsid=A250311135646686246",
                    "Origin" : "https://www.g2b.go.kr",
                    "priority" : "u=1, i",
                    "Referer" : "https://www.g2b.go.kr/",
                    "Sec-Ch-Ua" : '"Chromium";v="134", "Not:A-Brand";v="24", "Microsoft Edge";v="134"',
                    "Sec-Ch-Ua-Mobile" : "?0",
                    "Sec-Ch-Ua-Platform" : '"Windows"',
                    "Sec-Fetch-Dest" : "empty",
                    "Sec-Fetch-Mode" : "cors",
                    "Sec-Fetch-Site" : "same-origin",
                    "Submissionid" : "totalSrchList",
                    "User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36 Edg/133.0.0.0",
                    "Usr-Id" : "null"
                }

                data = { "dlSrchParamM" : {
                    "currentPage" : 1,
                    "recordCountPerPage" : "100",
                    "searchKeyword" : order,
                    "bizNm" : order,
                    "startBizYmd" : onemonth,
                    "endBizYmd" : heute,
                    "prcmBsneAreaCd" : "조070001 조070002 조070003 조070004 조070005",
                    "frcpYn" : "N",
                    "laseYn" : "N",
                    "rsrvYn" : "N"
                    }
                }

                response = requests.post(url, headers = params, data = json.dumps(data))
                
                # 목록 조회, 1페이지 당 최대 100개
                order_list = dict(json.loads(response.text))['dlTotalSrchL']
                
                # 총 개수/페이지 수/마지막 페이지 개수
                total = order_list[0]['totCnt']
                page_cycle = math.ceil(total/100)
                fin_num = total%100

                loop_count = 100
                category = ""
                name = ""
                biz_num = ""
                num_value2 = ""
                phase = ""
                input_date = ""
                gongo_com = ""
                suyo_com = ""
                up_date = ""
                fin = ""

                # 본격 크롤링
                for cycle in range(page_cycle):
                    if cycle == page_cycle - 1:
                        loop_count = fin_num
                    
                    data = { "dlSrchParamM" : {
                        "currentPage" : cycle + 1,
                        "recordCountPerPage" : "100",
                        "searchKeyword" : order,
                        "bizNm" : order,
                        "startBizYmd" : onemonth,
                        "endBizYmd" : heute,
                        "prcmBsneAreaCd" : "조070001 조070002 조070003 조070004 조070005",
                        "frcpYn" : "N",
                        "laseYn" : "N",
                        "rsrvYn" : "N"
                        }
                    }

                    response = requests.post(url, headers = params, data = json.dumps(data))

                    order_list = dict(json.loads(response.text))['dlTotalSrchL']
                    
                    for i in range(loop_count):
                        for key, value in order_list[i].items():
                            # 단계구분
                            if key == "untySrchSeNm":
                                phase = value

                            # 업무구분
                            if key == "prcmBsneAreaNm":
                                category = value

                            # 사업명
                            if key == "bizNm":
                                name = value.replace("&#41;", "").replace("&lt;/b&gt;", "").replace("&#40;", "").replace("&lt;b&gt;", "")
                                
                            # 사업번호
                            if key == "bizNo":
                                biz_num = value

                            # 사업번호 - 뒤에 번호, 세부페이지 접근용
                            if key == "bizOrd":
                                num_value2 = value

                            # 사업일자
                            if key == "inptDt":
                                input_date = str(value[:4] + "-" + value[4:6] + "-" + value[6:8])

                            # 공고기관
                            if key == "pbancInstUntyGrpNm":
                                gongo_com = value

                            # 수요기관
                            if key == "dmstUntyGrpNm":
                                suyo_com = value

                            # 공고일자
                            if key == "pbancPstgDt":
                                up_date = str(value[:4] + "-" + value[4:6] + "-" + value[6:8])

                            # 마감된 입찰공고인지 확인, 공란이 아니면 마감
                            if key == "pbancSrchItm01":
                                fin = value

                        if ("용역" in category) and ("입찰" in phase) and (fin == ""):
                            df.loc[num] = [order, category, name, biz_num, input_date, gongo_com, suyo_com, up_date]
                            num += 1

                    time.sleep(random.randint(0, 3))

                break

            time.sleep(random.randint(0, 3))

        except IndexError: # 결과값이 없는 경우
            pass

    df = df.drop_duplicates(subset = ['사업번호'])
    df = df.sort_values(by = ['키워드', '사업명', '사업일자'], ascending=[True, True, True])

    for j in range(len(df.index)): 
        worksheet.write(j + 2, 0, df.iloc[j]['키워드'], cell_bold)
        worksheet.write(j + 2, 1, df.iloc[j]['업무구분'])
        worksheet.write(j + 2, 2, df.iloc[j]['사업명'])
        worksheet.write(j + 2, 3, df.iloc[j]['사업번호'])
        worksheet.write(j + 2, 4, df.iloc[j]['사업일자'])
        worksheet.write(j + 2, 5, df.iloc[j]['공고기관'])
        worksheet.write(j + 2, 6, df.iloc[j]['수요기관'])
        worksheet.write(j + 2, 7, df.iloc[j]['공고일자'])

    worksheet.autofilter(f'A2:H{len(df.index)}')

    print('나라장터 crawl complete.')
    
def news():
    worksheet = workbook.add_worksheet('네이버 NEWS')

    worksheet.write('A1', '키워드', cell_bold)
    worksheet.write('B1', '원문보기', cell_bold)
    worksheet.write('C1', '제목', cell_bold)
    worksheet.write('D1', '설명', cell_bold)
    worksheet.write('E1', '신문사', cell_bold)
    worksheet.write('F1', '업로드 시간', cell_bold)

    worksheet.freeze_panes(1, 0)

    worksheet.set_column('A:A', 15)
    worksheet.set_column('C:C', 65)
    worksheet.set_column('D:D', 65)
    worksheet.set_column('E:E', 15)
    worksheet.set_column('F:F', 12)

    titlelist = []
    contentlist = []
    num = 0
    fake = ""
    refine = ""
    allow = 1

    df = pd.DataFrame(columns = ['키워드', '원문보기', '제목', '설명', '신문사', '업로드시간'])

    for word in wordlist:
        news_url = parse.urlparse(f"https://search.naver.com/search.naver?ssc=tab.news.all&query={word}")

        news_query = parse.parse_qs(news_url.query)

        news_url = f"https://search.naver.com/search.naver?{parse.urlencode(news_query, doseq=True)}&sm=tab_opt&sort=0&photo=0&field=0&pd=3&ds={naver_date}&de={naver_date}&docid=&related=0&mynews=0&office_type=0&office_section_code=0&news_office_checked=&nso=so%3Ar%2Cp%3Afrom{naver_another}to{naver_another}&is_sug_officeid=0"
        
        response = requests.get(news_url)

        soup = BeautifulSoup(response.text, 'lxml')

        try:     
            article_num = 1

            for page in range(10):
                news_url = parse.urlparse(f"https://search.naver.com/search.naver?ssc=tab.news.all&query={word}")

                news_query = parse.parse_qs(news_url.query)

                news_url = f"https://search.naver.com/search.naver?{parse.urlencode(news_query, doseq=True)}&sm=tab_pge&sort=0&photo=0&field=0&pd=3&ds={naver_date}&de={naver_date}&cluster_rank=45&mynews=0&office_type=0&office_section_code=0&news_office_checked=&nso=so:r,p:from{naver_another}to{naver_another},a:all&start={article_num}"

                response = requests.get(news_url)

                soup = BeautifulSoup(response.text, 'lxml')

                data = soup.find("ul", class_ = 'list_news _infinite_list').find_all("div", class_ = 'sds-comps-vertical-layout sds-comps-full-layout I6obO60yNcW8I32mDzvQ')

                for i in data:
                    try:
                        title = i.find("span", class_ = 'sds-comps-text sds-comps-text-ellipsis sds-comps-text-ellipsis-1 sds-comps-text-type-headline1').text.strip()
                        fake = re.sub(r" *\[(.+)\] *", '', title).strip()
                        fake = re.sub(r"\([^)]*\)", '', fake).strip()
                        fake = re.sub(r"[-=+,#/\?:^$.@*\"※~&▶■◆●◇△▲♥%ㆍ·｜!』\\‘’“”˝|\(\)\[\]\<\>`\'…⋯》【】]", ' ', fake).strip()
                        fake = re.sub('\s+', ' ', fake).strip()

                    except IndexError:
                        pass
                    
                    contents = i.find("span", class_ = 'sds-comps-text sds-comps-text-ellipsis sds-comps-text-ellipsis-3 sds-comps-text-type-body1')

                    try:
                        if contents == None:
                            contents = "정보없음"

                        else:
                            contents = contents.text.strip()

                        if contents.count('=') >= 1:
                            contents = re.sub(r'[=]+', '', contents)
                            contents = contents.strip()
                            
                    except IndexError:
                        contents = "정보없음"

                    refine = re.sub(r" *\[(.+)\] *", '', contents).strip()
                    refine = re.sub(r"\([^)]*\)", '', refine).strip()
                    refine = re.sub(r"[-=+,#/\?:^$.@*\"※~&▶■◆●◇△▲♥%ㆍ·｜!』\\‘’“”˝|\(\)\[\]\<\>`\'…⋯》【】]", ' ', refine).strip()
                    refine = re.sub('\s+', ' ', refine).strip()

                    try:
                        cite = i.find("span", class_ = 'info sds-comps-text sds-comps-text-ellipsis sds-comps-text-ellipsis-1 sds-comps-text-type-body2 sds-comps-text-weight-sm').text
                        
                    except AttributeError:
                        cite = i.find("span", class_= 'sds-comps-text sds-comps-text-ellipsis sds-comps-text-ellipsis-1 sds-comps-text-type-body2 sds-comps-text-weight-sm').text 
                    
                    news_date = i.find_all("span", class_ = 'sds-comps-text sds-comps-text-type-body2 sds-comps-text-weight-sm')

                    if len(news_date) == 1: # 일자 및 시간
                        news_date = news_date[0].text
                    elif len(news_date) == 2: # 일자 및 시간/네이버뉴스
                        if news_date[0].text == "네이버뉴스":
                            news_date = news_date[1].text
                        else:
                            news_date = news_date[0].text
                    elif len(news_date) == 3: #~면 ~단/일자 및 시간/네이버뉴스
                        if news_date[1].text == "네이버뉴스":
                            news_date = news_date[2].text
                        else:
                            news_date = news_date[1].text           
                    else:
                        news_date = news_date[0].text

                    link = i.find("a", class_ = 'rzROnhjF0RNNRoyDaO81 W035WwZVZIWyuG66e5iI').get('href')

                    dup_title = [i for i in titlelist if fake[:2] in i]
                    dup_contents = [j for j in contentlist if refine[:2] in j]

                    if len(dup_title) >= 1:
                        for j in dup_title:
                            if int(f'{SequenceMatcher(None, fake, j).ratio()*100:.0f}') >= 45:
                                # print('fake dup list', fake, int(f'{SequenceMatcher(None, fake, j).ratio()*100:.0f}'))
                                allow = 0
                    
                    if len(dup_contents) >= 1:
                        for k in dup_contents:
                            if int(f'{SequenceMatcher(None, refine, k).ratio()*100:.0f}') >= 45:
                                # print('refine dup list', refine, int(f'{SequenceMatcher(None, refine, k).ratio()*100:.0f}'))
                                allow = 0

                    if (word in title) and (allow == 1):
                        titlelist.append(fake)
                        contentlist.append(refine)

                        df.loc[num] = [word, link, title, contents, cite, news_date]
                        num += 1
                        
                    allow = 1

                time.sleep(random.randint(0, 3))
            
                article_num += 10

        except TypeError as te:
            logging.info(f'[{word}]TypeError: {te}')
            pass

        except AttributeError as ae: # 검색결과가 없는 경우
            logging.info(f'[{word}]AttributeError: {ae}')
            pass

    df = df.sort_values(by = '제목')

    for j in range(len(df.index)): 
        worksheet.write(j + 1, 0, df.iloc[j]['키워드'])
        worksheet.write(j + 1, 1, df.iloc[j]['원문보기'])
        worksheet.write(j + 1, 2, df.iloc[j]['제목'])
        worksheet.write(j + 1, 3, df.iloc[j]['설명'])
        worksheet.write(j + 1, 4, df.iloc[j]['신문사'])
        worksheet.write(j + 1, 5, str(df.iloc[j]['업로드시간']))

    worksheet.autofilter(f'A1:F{len(df.index)}')

    logging.info(f'NEWS crawl complete.')
    print('NEWS crawl complete.')

def view():
    worksheet = workbook.add_worksheet('네이버 VIEW')

    worksheet.write('A1', '키워드', cell_bold)
    worksheet.write('B1', '원문보기', cell_bold)
    worksheet.write('C1', '제목', cell_bold)
    worksheet.write('D1', '출처', cell_bold)
    worksheet.write('E1', '날짜', cell_bold)
    worksheet.write('F1', '비고', cell_bold)

    worksheet.freeze_panes(1, 0) 

    worksheet.set_column('A:A', 15)
    worksheet.set_column('C:C', 70)
    worksheet.set_column('D:D', 40)
    worksheet.set_column('E:E', 12)

    allow = 1
    titlelist = []
    num = 0

    df = pd.DataFrame(columns = ['키워드', '원문보기', '제목', '출처', '날짜', '비고'])

    for word in wordlist:

        view_url = parse.urlparse(f"https://search.naver.com/search.naver?where=view&sm=tab_jum&query={word}")

        view_query = parse.parse_qs(view_url.query)

        view_url = f"https://search.naver.com/search.naver?{parse.urlencode(view_query, doseq=True)}&nso=so%3Ar%2Cp%3Afrom{naver_another}to{naver_another}%2Ca%3Aall"

        response = requests.get(view_url)

        soup = BeautifulSoup(response.text, 'lxml')
        data = soup.find_all("li", class_ = 'bx _svp_item')
   
        for i in data:
            title = i.find("a", class_ = 'api_txt_lines total_tit _cross_trigger').text.strip()
            cite = i.find("a", class_ = 'sub_txt sub_name').text 
            view_date = i.find("span", class_ = 'sub_time sub_txt').text
            link = i.find("a", class_ = 'api_txt_lines total_tit _cross_trigger').get('href')

            try:
                if cite[0] == "=":
                    cite = cite[1:]
            except IndexError:
                pass

            dup = [i for i in titlelist if title[:2] in i]

            if len(dup) >= 1:
                for j in dup:
                    if int(f'{SequenceMatcher(None, title, j).ratio()*100:.0f}') >= 40:
                        allow = 0

            if (word in title) and (allow == 1):
                titlelist.append(title)
                df.loc[num] = [word, link, title, cite, view_date, '']
                num += 1

            allow = 1
        
        time.sleep(1)

        # 네이버 동영상
        video_url = parse.urlparse(f"https://search.naver.com/search.naver?where=video&sm=tab_jum&query={word}")

        video_query = parse.parse_qs(video_url.query)

        video_url = f"https://search.naver.com/search.naver?{parse.urlencode(video_query, doseq=True)}&nso=so%3Ar%2Cp%3Afrom{naver_another}to{naver_another}%2Ca%3Aall"

        response = requests.get(video_url)

        soup = BeautifulSoup(response.text, 'lxml')

        data = soup.find_all("li", class_ = 'video_item _svp_item')
    
        for i in data:
            title = i.find("a", class_ = 'info_title').text.strip()

            if i.find("a", class_ = 'channel') == None:
                cite = "정보없음"
            else:
                cite = i.find("a", class_ = 'channel').text 

            if len(i.find_all("span", class_ = 'desc')) >= 2:
                video_date = i.find_all("span", class_ = 'desc')[1].text 
            else:
                video_date = i.find("span", class_ = 'desc').text 

            link = i.find("a", class_ = 'info_title').get('href')

            dup = [i for i in titlelist if title[:2] in i]
        
            if len(dup) >= 1:
                for j in dup:
                    if int(f'{SequenceMatcher(None, title, j).ratio()*100:.0f}') >= 40:
                        allow = 0

            if (word in title) and (allow == 1):
                titlelist.append(title)
                df.loc[num] = [word, link, title, cite, video_date, '동영상']
                num += 1
        
            allow = 1
        
        time.sleep(random.randint(0, 3))

    df = df.drop_duplicates(subset = ['제목'])
    df = df.sort_values(by = '제목')

    for i in range(len(df.index)):
        worksheet.write(i + 1, 0, df.iloc[i]['키워드'])
        worksheet.insert_image(i + 1, 1, 'go.PNG', {'url': df.iloc[i]['원문보기'], 'x_scale': 0.06, 'y_scale': 0.06, 'x_offset': 25})
        
        try:
            if df.iloc[i]['출처'][0] != '':
                if df.iloc[i]['출처'][0] == '=':
                    worksheet.write(i + 1, 3, df.iloc[i]['출처'][1:])

                else:
                    worksheet.write(i + 1, 3, df.iloc[i]['출처'])

            else:
                df.iloc[i]['출처'] = '없음'
        except IndexError:
                df.iloc[i]['출처'] = '없음'

        if df.iloc[i]['제목'][0] != '':
            if df.iloc[i]['제목'][0] == '=':
                worksheet.write(i + 1, 2, df.iloc[i]['제목'][1:])

            else:
                worksheet.write(i + 1, 2, df.iloc[i]['제목'])

        else:
            df.iloc[i]['제목'] = '없음'

        worksheet.write(i + 1, 4, df.iloc[i]['날짜'])
        worksheet.write(i + 1, 5, df.iloc[i]['비고'], cell_bold)

    worksheet.autofilter(f'A1:F{len(df.index)}')

    logging.info(f'VIEW and VIDEO crawl complete.')
    print('VIEW and VIDEO crawl complete.')

    workbook.close()

def email(receiver):
    message = MIMEMultipart()

    message['Subject'] = '나라장터 관련 키워드 네이버 검색결과'
    message['From'] = formataddr(('RA팀', "dailyinfo@pkf.kr"))
    message['To'] = formataddr(receiver)

    template = Template("""<html>
                                    <head></head>
                                    <body>
                                        안녕하세요.<br>
                                        금일 검색된 결과는 첨부파일과 같습니다.<br><br>
                                        대상 키워드는<br>
                                        <b>자체감사, 외부감사, 회계감사, 감사시스템, 내부감사, 내부회계, 모니터링, 회계감리, 회계법인, 회계정산,<br>
                                        위탁정산, IT감사, 전산감사, ESG, 내부통제, 감사위원, 지배구조, 위험관리, e감사, 품질평가, 연구개발혁신법,<br>
                                        회계연도, 법인세, 국세, 지방세, 세무, 전략, 감사체계, 감사품질, 직무평가, 감사정보, 감사행정, 감사모니터링, 정보보호</b><br>
                                        총 34개 입니다.<br><br>
                                        감사합니다.<br>
                                        RA팀 드림.
                                    </body>
                                </html>""")

    str_msg  = template.safe_substitute() 
    mime_msg = MIMEText(str_msg, 'html')
    message.attach(mime_msg)

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(f"{filename}", "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={filename}')
    message.attach(part)

    with smtplib.SMTP_SSL('mail.pkf.kr', 465) as server:
        server.ehlo()
        server.login('dailyinfo@pkf.kr', '1111')
        server.send_message(message)
    
    logging.info(f'{receiver}')
    print(f'{receiver} Message sent complete.')

def email_to_admin(receiver):
    message = MIMEMultipart()

    message['Subject'] = 'crawl.py 이메일 발송 오류'
    message['From'] = formataddr(("RA팀", "dailyinfo@pkf.kr"))
    message['To'] = formataddr(("우혜진", "haejin.woo@pkf.kr"))

    template = Template(f"<html><body> {receiver} 에서 발송 멈춤. </body></html>")

    str_msg  = template.safe_substitute() 
    mime_msg = MIMEText(str_msg, 'html')
    message.attach(mime_msg)

    with smtplib.SMTP_SSL('mail.pkf.kr', 465) as server:
        server.ehlo()
        server.login('dailyinfo@pkf.kr', '1111')
        server.send_message(message)

    logging.info("Message sent to admin.")

if __name__ == "__main__":
    list1 = [('올파트너스', 'sh_allpartners@pkf.kr'), ('신동복', 'dongbok.shin@pkf.kr'), ('김학수', 'haksoo.kim@pkf.kr'), ('최준기', 'juneki.choi@pkf.kr'), ('오영주', 'youngju.oh@pkf.kr'), ('윤영광', 'youngkwang.yoon@pkf.kr'), ('이예원', 'yewon.lee763@pkf.kr'),\
            ('박성래', 'sungrae.park@pkf.kr'), ('이재덕', 'jaeduk.lee@pkf.kr'), ('최종규', 'jonggyu.choi@pkf.kr'), ('신용선', 'yongsun.shin@pkf.kr'), ('이재인', 'jaein.lee@pkf.kr'), ('이한비', 'hanbi.lee@pkf.kr')]
    
    g2b()
    news()
    view()
    
    for i in list1:
        try:
            email(i)
            time.sleep(1)

        except smtplib.SMTPRecipientsRefused:
            logging.error(f'Email Refused in {i}')
            email_to_admin(i)
            pass
        
        except TimeoutError:
            logging.error(f'Email Timeout Error in {i}')
            email_to_admin(i)
            pass

        except smtplib.SMTPServerDisconnected:
            logging.error(f'Email Disconnected Error in {i}')
            email_to_admin(i)
            pass