# -*- coding: utf-8 -*-
# 정보보호 공시 조회 사이트는 원칙적으로 크롤링 금지 사이트임
# https://stackoverflow.com/questions/59182694/pdfminer-extraction-for-single-words-lttext-lttextbox
import requests
import xlsxwriter
import time
import os
import random
import pytesseract
from pdf2image import convert_from_path
import glob

from bs4 import BeautifulSoup
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter, PDFPageAggregator
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFPage
from io import StringIO, open
from requests import get
from pdfminer.layout import LAParams, LTText
from PyPDF2 import PdfFileWriter, PdfFileReader

keyword = ['감리법인', '감 리 법 인', '회계법인', '회 계 법 인']

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

    pdfs = glob.glob(rf'C:\Users\user\Desktop\작업폴더\정보보호공시\2024\{filename}')

    for pdf_path in pdfs:
        pages = convert_from_path(pdf_path, 500, poppler_path = r'C:\Program Files\poppler-23.07.0\Library\bin')

        pytesseract.pytesseract.tesseract_cmd = r'C:\Users\user\AppData\Local\Programs\Tesseract-OCR\tesseract'

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
                break

    return gamri

def read_file3(filename):
    gamri = "직접 확인 필요"

    pdfs = glob.glob(rf'C:\Users\user\Desktop\작업폴더\정보보호공시\2024\{filename}')

    for pdf_path in pdfs:
        pages = convert_from_path(pdf_path, 500, poppler_path = r'C:\Program Files\poppler-23.07.0\Library\bin')

        pytesseract.pytesseract.tesseract_cmd = r'C:\Users\user\AppData\Local\Programs\Tesseract-OCR\tesseract'

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
                break

    return gamri               

# 요거로 수행함
def start_here2():
    workbook = xlsxwriter.Workbook('2024년 정보보호 공시 기업현황1.xlsx')

    cell_center = workbook.add_format({'align': 'center'})
    cell_bold = workbook.add_format({'bold': True, 'align': 'center'})
    cell_bold.set_bg_color('#dce6f1')
    cell_number = workbook.add_format({'num_format': '#,##0', 'align': 'center'})

    worksheet = workbook.add_worksheet('기업 목록')

    worksheet.write('A1', '게시일', cell_bold)
    worksheet.write('B1', '기업명', cell_bold)
    worksheet.write('C1', '업종', cell_bold)
    worksheet.write('D1', '최종수정일', cell_bold)
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
    worksheet.write('O1', '의무/자율', cell_bold)

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
    worksheet.set_column('N:N', 12)
    worksheet.set_column('O:O', 10)

    count = 0
    row_num = 1

    # 742
    response = requests.get(f"https://isds.kisa.or.kr/kr/publish/list.do?menuNo=204942&pageNum=1&limit=742")

    soup = BeautifulSoup(response.text, 'lxml')

    url_list = soup.find('tbody').find_all('tr')

    for i in url_list:
        count += 1
        prefile = "X"
        postfile = "X"
        test_company = "직접 확인 필요"
        prefile_num = 0

        try:
            upload_date = i.find_all('td')[5].text.strip()

            force = i.find_all('td')[2].text.strip()

            company_num = int(i.find_all('td')[3].find('a').get('onclick').split('(')[1].split(',')[0])

            response = requests.get(f"https://isds.kisa.or.kr/kr/publish/view.do?menuNo=204942&publishNo={company_num}")

            soup = BeautifulSoup(response.text, 'lxml')

            # 기업정보
            info = soup.find_all('td')

            company = info[2].text.strip()
            category = info[3].text.strip()

            check_postfile = info[5].text.strip()      

            if len(check_postfile) != 0:
                postfile = "O"

            check_prefile = info[6].text.strip()       
            
            if len(check_prefile) != 0:
                prefile = "O"

                prefile_num = info[6].find('a', href=True)['href'].strip().split('(')[1].split(')')[0]

                # 파일 다운로드
                with open(check_prefile, "wb") as file:
                    response = requests.get(f"https://isds.kisa.or.kr/kr/user/publish/fileDownload.do?attachNo={prefile_num}")
                
                    file.write(response.content)
                
                # 첫번째 시도
                # try:
                #     test_company = read_file(check_prefile)
                # except:
                #     print('Error in First')

                # 두번째 시도
                if test_company == "직접 확인 필요":
                    try:
                        test_company = read_file2(check_prefile)
                    except:
                        print('Error in Second')

                # 세번째 시도
                if test_company == "직접 확인 필요":
                    try:
                        test_company = read_file3(check_prefile)
                    except:
                        print('Error in Third')
                        test_company = "직접 확인 필요"

            else:
                prefile = "X"
                test_company = "해당없음"


            print(company, force, category, postfile, prefile, prefile_num, test_company)

            # 정보보호 현황
            info2 = soup.find_all('td')[6:]
            last_date = str(soup.find_all('span', class_ = 'date')[0])[-21:-11]

            try:
                it_cost = info2[1].text.strip()[:-1].replace(",", "")
                security_cost = info2[2].text.strip()[:-1].replace(",", "")
                
                if it_cost == "":
                    it_cost = "0"

                if security_cost == "":
                    security_cost = "0"

            except ValueError as ve:
                it_cost = "0"
                security_cost = "0"

            except IndexError as ie:
                it_cost = "0"
                security_cost = "0"

            cost_rate = info2[4].text.strip()[:-1]

            if cost_rate == "":
                cost_rate = "0"

            it_people = info2[7].text.strip()[:-1]

            if it_people == "":
                it_people = "0"

            try:
                security_people = info2[8].find_all('dd')[0].text.strip()[:-1]

                if security_people == "":
                    security_people = "0"

                inner_people = info2[8].find_all('dd')[1].text.strip()[:-1]

                if inner_people == "":
                    inner_people = "0"

                outer_people = info2[8].find_all('dd')[2].text.strip()[:-1]

                if outer_people == "":
                    outer_people = "0"
                
            except IndexError as ie:
                security_people = "0"
                inner_people = "0"
                outer_people = "0"

            people_rate = info2[9].text.strip()[:-1]

            if people_rate == "":
                people_rate = "0"

            worksheet.write(row_num, 0, upload_date, cell_center)
            worksheet.write(row_num, 1, company)
            worksheet.write(row_num, 2, category)
            worksheet.write(row_num, 3, last_date, cell_center)
            worksheet.write(row_num, 3, prefile, cell_center)
            worksheet.write(row_num, 4, postfile, cell_center)
            worksheet.write(row_num, 5, test_company)
            worksheet.write(row_num, 6, it_cost.strip(), cell_number)
            worksheet.write(row_num, 7, security_cost.strip(), cell_number)
            worksheet.write(row_num, 8, cost_rate.strip(), cell_center)
            worksheet.write(row_num, 9, it_people.strip(), cell_center)
            worksheet.write(row_num, 10, security_people.strip(), cell_center)
            worksheet.write(row_num, 11, inner_people.strip(), cell_center)
            worksheet.write(row_num, 12, outer_people.strip(), cell_center)
            worksheet.write(row_num, 13, people_rate.strip(), cell_center)
            worksheet.write(row_num, 14, force, cell_center)

            row_num += 1

            time.sleep(random.randrange(1,3))

        except ConnectionError as ce:
            print(f'ConnectionError in {company}, {ce}')
            pass
            
        except ValueError as ve:
            print(f'ValueError in {company}, {ve}')
            pass

        print(count)

    workbook.close()

def extract_test_company():
    workbook = xlsxwriter.Workbook('사전점검확인서 발행 기업 목록.xlsx')

    cell_bold = workbook.add_format({'bold': True, 'align': 'center'})

    worksheet = workbook.add_worksheet('기업 목록')

    worksheet.write('A1', '기업명', cell_bold)
    worksheet.write('B1', '사전점검확인서', cell_bold)
    worksheet.write('C1', '감리기업명', cell_bold)
    
    response = requests.get(f"https://isds.kisa.or.kr/publish/list.do?pageNum=1&limit=650")

    soup = BeautifulSoup(response.text, 'lxml')

    url_list = soup.find_all('tr')

    row_num = 0
    test_company = "직접 확인 필요"
    prefile = "X"

    pytesseract.pytesseract.tesseract_cmd = r'C:\Users\user\AppData\Local\Programs\Tesseract-OCR\tesseract'

    for i in url_list:
        test_company = "직접 확인 필요"
        prefile = "X"
        
        company_num = int(i.find_all('td')[2].find('a').get('onclick').split('(')[1].split(',')[0])

        response = requests.get(f"https://isds.kisa.or.kr/publish/view.do?publishNo={company_num}")

        soup = BeautifulSoup(response.text, 'lxml')

        info = soup.find_all('div', class_ = 'form-area')[0].find_all('tr')
        company = info[1].find('input', {'id': 'corpName'}).get('value')
        check_prefile = info[5].find('input', {'id': 'fileName2'}).get('value')

        if len(check_prefile) != 0:
            prefile = "O"

            # 첫번째 시도
            try:
                test_company = read_file(check_prefile)
            except:
                print('Error in First')

            # 두번째 시도
            if test_company == "직접 확인 필요":
                try:
                    test_company = read_file2(check_prefile)
                except:
                    print('Error in Second')
                    test_company = read_file3(check_prefile)

            # 세번째 시도
            if test_company == "직접 확인 필요":
                try:
                    test_company = read_file3(check_prefile)
                except:
                    print('Error in Third')
                    test_company = "직접 확인 필요"
        else:
            prefile = "X"
            test_company = "해당없음"

        if prefile == "O":
            print(company, test_company)

        worksheet.write(row_num, 0, company)
        worksheet.write(row_num, 1, prefile)
        worksheet.write(row_num, 2, test_company)

        row_num += 1

    workbook.close()

if __name__ == "__main__":
    start_here2()
    # extract_test_company()