from bs4 import BeautifulSoup
import requests, openpyxl

res = requests.get('http://corners.gmarket.co.kr/Bestsellers?viewType=G&groupCode=G06')
soup = BeautifulSoup(res.content, 'html.parser')

itemlist = soup.select('div.best-list')
item = itemlist[1]
product = item.select('ul > li')

# 엑셀 파일 만들기
excel_file = openpyxl.Workbook()
excel_sheet = excel_file.active
excel_sheet.title = 'best_product'

excel_sheet.append(['랭킹', '상품명', '판매가격', '판매업체', '상품상세링크'])
excel_sheet.column_dimensions['A'].width = 5
excel_sheet.column_dimensions['B'].width = 70
excel_sheet.column_dimensions['C'].width = 15
excel_sheet.column_dimensions['D'].width = 20
excel_sheet.column_dimensions['E'].width = 80
# 엑셀파일 행 이름 가운데정렬
# 엑셀파일에서 인덱스는 0이 아닌 1부터 시작
for i in range(5):
    excel_sheet.cell(row=1, column=i+1).alignment = openpyxl.styles.Alignment(horizontal='center') 
   

    
for i, data in enumerate(product[:10]):    # 리스트 슬라이싱 해서 10개만 나오게 하기
    dataname = data.select_one('a.itemname')    
    
    dataprice = data.select_one('div.item_price span > span')
    
    res_2 = requests.get(dataname['href'])
    soup_2 = BeautifulSoup(res_2.content, 'html.parser')
    datacompany = soup_2.select_one('#container > div.item-topinfowrap span.text__seller > a')
    
    print(i+1, dataname.get_text(), dataprice.get_text(), datacompany.get_text(), dataname['href'])
    
    # 엑셀파일에 데이터 추가
    excel_sheet.append([i+1, dataname.get_text(), dataprice.get_text(), datacompany.get_text(), dataname['href']])
    # 엑셀파일 데이터에 하이퍼링크 걸기
    excel_sheet.cell(row=i+2 , column=5).hyperlink = dataname['href']
     # 엑셀파일 글자 중앙정렬
    excel_sheet.cell(row=i+2, column=1).alignment = openpyxl.styles.Alignment(horizontal='center')
    # 엑셀파일 글자 색 바꾸기
    excel_sheet.cell(row=i+2, column=2).font = openpyxl.styles.Font(color="01579B")

    
# 엑셀 파일 저장
excel_file.save('gmarket.xlsx')
excel_file.close()
