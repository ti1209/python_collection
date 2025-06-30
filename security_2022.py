# -*- coding: utf-8 -*-
# 정보보호 공시 조회 사이트는 원칙적으로 크롤링 금지 사이트임
# https://stackoverflow.com/questions/59182694/pdfminer-extraction-for-single-words-lttext-lttextbox

import requests
import xlsxwriter
import time
import os
import random
# import pytesseract
# from pdf2image import convert_from_path
# import glob

from bs4 import BeautifulSoup
# from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
# from pdfminer.converter import TextConverter, PDFPageAggregator
# from pdfminer.pdfpage import PDFPage
# from pdfminer.pdfpage import PDFPage
from io import StringIO, open
from requests import get
# from pdfminer.layout import LAParams, LTText
# from PyPDF2 import PdfFileWriter, PdfFileReader

keyword = ['감리법인', '감 리 법 인', '회계법인', '회 계 법 인']

# 첨부파일 다운로드
def download_file(filename, url):
    with open(filename, "wb") as file:
        response = get('https://www.ksecurity.or.kr' + url)

        file.write(response.content)       

# pdf 읽기
def read_file(filename):
    gamri = "직접 확인 필요"

    fp = open(filename, 'rb')

    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    dev = PDFPageAggregator(rsrcmgr, laparams=laparams)
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, dev)

    pages = PDFPage.get_pages(fp)

    for pageNumber, page in enumerate(PDFPage.get_pages(fp)):
        if pageNumber == 1: # 2페이지에 감리기업 정보가 있음
            interpreter.process_page(page)
            layout = dev.get_result()

            for lobj in layout:
                if isinstance(lobj, LTText):
                    x, y, text = lobj.bbox[0], lobj.bbox[3], lobj.get_text()
                    
                    for word in keyword:
                        if word in lobj.get_text():
                            # print('At %r is text: %s' % ((x, y), text.strip()))
                            gamri = text.strip()

    fp.close()

    return gamri

def read_file2(filename):
    gamri = "직접 확인 필요"

    pdfs = glob.glob(rf'C:\Users\이현세무법인\Desktop\작업폴더\회사별 자료\2022\정보보호 공시\사전점검확인서\{filename}')

    for pdf_path in pdfs:
        pages = convert_from_path(pdf_path, 500, poppler_path = r'C:\Program Files\poppler-22.04.0\Library\bin')

        # pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'

        for pageNum,imgBlob in enumerate(pages):
            custom_config = r'--oem 3 --psm 6'
            text = pytesseract.image_to_string(imgBlob, config = custom_config, lang = 'kor')
            
            if pageNum == 1:
                with open(f'{pdf_path[:-4]}_page{pageNum}.txt', 'w') as the_file:
                    the_file.write(text)
                
                with open(f'{pdf_path[:-4]}_page{pageNum}.txt', "r") as file:
                    line = file.readlines()[::-1]

                    for i in line:
                        for word in keyword:
                            if word in i:
                                gamri = i.strip()
                                break

    return gamri

def read_file3(filename):
    gamri = "직접 확인 필요"

    pdfs = glob.glob(rf'C:\Users\이현세무법인\Desktop\작업폴더\회사별 자료\2022\정보보호 공시\사전점검확인서\{filename}')

    for pdf_path in pdfs:
        pages = convert_from_path(pdf_path, 500, poppler_path = r'C:\Program Files\poppler-22.04.0\Library\bin')

        # pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'

        for pageNum,imgBlob in enumerate(pages):
            custom_config = r'--oem 3 --psm 6'
            text = pytesseract.image_to_string(imgBlob, config = custom_config, lang = 'kor')
            
            if pageNum == 0:
                with open(f'{pdf_path[:-4]}_page{pageNum}.txt', 'w') as the_file:
                    the_file.write(text)
                
                with open(f'{pdf_path[:-4]}_page{pageNum}.txt', "r") as file:
                    line = file.readlines()[::-1]

                    for i in line:
                        for word in keyword:
                            if word in i:
                                gamri = i.strip()
                                break

    return gamri

def start_here():
    workbook = xlsxwriter.Workbook('정보보호 공시 기업 목록_220704.xlsx')

    cell_center = workbook.add_format({'align': 'center'})
    cell_bold = workbook.add_format({'bold': True})

    worksheet = workbook.add_worksheet('기업 목록')

    worksheet.write('A1', '공시연도', cell_bold)
    worksheet.write('B1', '게시일', cell_bold)
    worksheet.write('C1', '기업명', cell_bold)
    worksheet.write('D1', '사전점검확인서', cell_bold)
    worksheet.write('E1', '감리기업명', cell_bold)

    worksheet.freeze_panes(1, 0)

    worksheet.set_column('B:B', 11)
    worksheet.set_column('C:C', 24)
    worksheet.set_column('D:D', 14)

    row_num = 1
    prefile = "X"

    time.sleep(3) # 크롤링 봇 탐지 방지를 위한 sleep

    # 총 페이지 수 알아내기
    url = "https://www.ksecurity.or.kr/user/extra/kisis/34/disclosure/disclosureList/jsp/LayOutPage.do"

    header = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
        'Cookie': 'JSESSIONID=BC97D72AC17A6270BFD36D5FD6374C9A',
        'Host': 'www.ksecurity.or.kr',
        'Referer': 'https://www.ksecurity.or.kr/kisis/subIndex/33.do',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.82 Safari/537.36'
    }

    response = requests.get(url, headers = header)

    soup = BeautifulSoup(response.text, 'lxml')

    total_page = int(soup.find("div", class_ = 'floatL').text[-6:-4])

    # 본격적으로 데이터 가져오기
    for page in range(total_page):
        url = f"https://www.ksecurity.or.kr/user/extra/kisis/34/disclosure/disclosureList/jsp/LayOutPage.do?column=&search=&spage={page + 1}"

        header = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
            'Cache-Control': 'max-age=0',
            'Connection': 'keep-alive',
            'Cookie': 'JSESSIONID=BC97D72AC17A6270BFD36D5FD6374C9A',
            'Host': 'www.ksecurity.or.kr',
            'Referer': 'https://www.ksecurity.or.kr/kisis/subIndex/33.do',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.82 Safari/537.36'
        }

        response = requests.get(url, headers = header)

        soup = BeautifulSoup(response.text, 'lxml')

        data_table = soup.find("table", class_ = 'table_style01 mTs').find('tbody').find_all('tr')

        for i in data_table:
            publish_year = i.find_all('td')[1].text
            company = i.find_all('td')[2].text.strip()
            company_url = i.find_all('td')[2].find('a').get('href')
            upload_date = i.find_all('td')[3].text

            # 각 상세페이지로 가서 첨부파일 다운로드
            detail_url = f"https://www.ksecurity.or.kr{company_url}"

            header = {
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                'Accept-Encoding': 'gzip, deflate, br',
                'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
                'Cache-Control': 'max-age=0',
                'Connection': 'keep-alive',
                'Cookie': 'JSESSIONID=BC97D72AC17A6270BFD36D5FD6374C9A',
                'Host': 'www.ksecurity.or.kr',
                'Referer': 'https://www.ksecurity.or.kr/kisis/subIndex/33.do',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.82 Safari/537.36'
            }

            response = requests.get(detail_url, headers = header)

            soup = BeautifulSoup(response.text, 'lxml')

            # 첨부파일2가 첨부되어있는지 확인(첨부파일2의 이름 길이)
            check_file = soup.find('table', class_ = 'table_style02 mTs').find_all('tr')[2].text[6:].strip()

            if len(check_file) != 0:
                prefile = "O"

                file_url = soup.find('table', class_ = 'table_style02 mTs').find_all('tr')[2].find('a').get('href')

                download_file(check_file, file_url)
                gamri_company = read_file(check_file)
                
                try:
                    os.rename(check_file, check_file[:-4] + '_' + publish_year[:-1] + '.pdf')
                except FileExistsError:
                    print(company, check_file)

                # print(company, gamri_company)
            else:
                prefile = "X"
                gamri_company = "해당없음"

            worksheet.write(row_num, 0, publish_year,cell_center)
            worksheet.write(row_num, 1, upload_date, cell_center)
            worksheet.write(row_num, 2, company)
            worksheet.write(row_num, 3, prefile, cell_center)
            worksheet.write(row_num, 4, gamri_company)

            row_num += 1

    workbook.close()

def start_here2():
    workbook = xlsxwriter.Workbook('2022년 정보보호 공시 기업현황.xlsx')

    cell_center = workbook.add_format({'align': 'center'})
    cell_bold = workbook.add_format({'bold': True, 'align': 'center'})
    cell_bold.set_bg_color('#dce6f1')
    cell_number = workbook.add_format({'num_format': '#,##0', 'align': 'center'})

    worksheet = workbook.add_worksheet('기업 목록')

    worksheet.write('A1', '게시일', cell_bold)
    worksheet.write('B1', '기업명', cell_bold)
    worksheet.write('C1', '업종', cell_bold)
    worksheet.write('D1', '사전점검확인서', cell_bold)
    worksheet.write('E1', '사후검증동의서', cell_bold)
    worksheet.write('F1', '감리기업명', cell_bold)
    worksheet.write('G1', '정보기술 부문 투자액(원)', cell_bold)
    worksheet.write('H1', '정보보호 부문 투자액(원)', cell_bold)
    worksheet.write('I1', '투자액 비율(%)', cell_bold)
    worksheet.write('J1', '정보기술 부문 인력(명)', cell_bold)
    worksheet.write('K1', '정보보호 전담인력(명)', cell_bold)
    worksheet.write('L1', '정보보호 내부인력(명)', cell_bold)
    worksheet.write('M1', '정보보호 외부인력(명)', cell_bold)
    worksheet.write('N1', '인력 비율(%)', cell_bold)

    worksheet.freeze_panes(1, 0)

    worksheet.set_column('A:A', 11)
    worksheet.set_column('B:B', 24)
    worksheet.set_column('C:C', 24)
    worksheet.set_column('D:D', 14)
    worksheet.set_column('E:E', 14)
    worksheet.set_column('F:F', 24)
    worksheet.set_column('G:G', 22)
    worksheet.set_column('H:H', 22)
    worksheet.set_column('I:I', 13)
    worksheet.set_column('J:J', 20)
    worksheet.set_column('K:K', 20)
    worksheet.set_column('L:L', 20)
    worksheet.set_column('M:M', 20)
    worksheet.set_column('N:N', 12)

    count = 0
    row_num = 1

    response = requests.get(f"https://isds.kisa.or.kr/kr/publish/list.do?menuNo=204942&pageNum=1&limit=713")

    soup = BeautifulSoup(response.text, 'lxml')

    url_list = soup.find_all('tr')

    # pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'

    for i in url_list:
        count += 1
        prefile = "X"
        postfile = "X"
        test_company = "직접 확인 필요"
        prefile_num = 0

        try:
            upload_date = i.find_all('td')[4].text.strip()

            company_num = int(i.find_all('td')[2].find('a').get('onclick').split('(')[1].split(',')[0])
            print(company_num)
            response = requests.get(f"https://isds.kisa.or.kr/publish/view.do?publishNo={company_num}")

            soup = BeautifulSoup(response.text, 'lxml')

            # 기업정보
            info = soup.find_all('div', class_ = 'form-area')[0].find_all('tr')

            company = info[1].find('input', {'id': 'corpName'}).get('value')
            category = info[2].find('span').text
            check_postfile = info[4].find('input', {'id': 'fileName2'}).get('value')       

            if len(check_postfile) != 0:
                postfile = "O"

            check_prefile = info[5].find('input', {'id': 'fileName2'}).get('value')       

            if len(check_prefile) != 0:
                prefile = "O"

                prefile_num = int(info[5].find('label').get('onclick').split('(')[1].split(')')[0])

                # 파일 다운로드
                # with open(check_prefile, "wb") as file:
                #     response = requests.get(f"https://isds.kisa.or.kr/publish/fileDownload.do?attachNo={prefile_num}")
                
                #     file.write(response.content)       

                # # 첫번째 시도
                # try:
                #     test_company = read_file(check_prefile)
                # except:
                #     print('Error in First')

                # # 두번째 시도
                # if test_company == "직접 확인 필요":
                #     try:
                #         test_company = read_file2(check_prefile)
                #     except:
                #         print('Error in Second')
                #         test_company = read_file3(check_prefile)

                # # 세번째 시도
                # if test_company == "직접 확인 필요":
                #     try:
                #         test_company = read_file3(check_prefile)
                #     except:
                #         print('Error in Third')
                #         test_company = "직접 확인 필요"
                
                # print(test_company)

            else:
                prefile = "X"
                test_company = "해당없음"

            # 정보보호 현황
            info2 = soup.find_all('div', class_ = 'form-area')[1]

            it_cost = info2.find('input', {'id': 'investAmountA'}).get('value')
            security_cost = info2.find('input', {'id': 'investAmountB'}).get('value')
            cost_rate = info2.find('input', {'id': 'investRatio'}).get('value')
            it_people = info2.find('input', {'id': 'hrIt'}).get('value')
            security_people = info2.find('input', {'id': 'hrItTotal'}).get('value')
            inner_people = info2.find('input', {'id': 'hrItIn'}).get('value')
            outer_people = info2.find('input', {'id': 'hrItOut'}).get('value')
            people_rate = info2.find('input', {'id': 'hrRatio'}).get('value')

            worksheet.write(row_num, 0, upload_date, cell_center)
            worksheet.write(row_num, 1, company)
            worksheet.write(row_num, 2, category)
            worksheet.write(row_num, 3, prefile, cell_center)
            worksheet.write(row_num, 4, postfile, cell_center)
            worksheet.write(row_num, 5, test_company)
            worksheet.write(row_num, 6, int(it_cost), cell_number)
            worksheet.write(row_num, 7, int(security_cost), cell_number)
            worksheet.write(row_num, 8, float(cost_rate), cell_center)
            worksheet.write(row_num, 9, float(it_people), cell_center)
            worksheet.write(row_num, 10, float(security_people), cell_center)
            worksheet.write(row_num, 11, float(inner_people), cell_center)
            worksheet.write(row_num, 12, float(outer_people), cell_center)
            worksheet.write(row_num, 13, float(people_rate), cell_center)

            row_num += 1

            time.sleep(random.randrange(1,3))

        except ConnectionError:
            print(f'ConnectionError in {company}')
            pass
            
        except ValueError:
            print(f'ValueError in {company}')

        print(count)
        
    workbook.close()

if __name__ == "__main__":
    start_here2()