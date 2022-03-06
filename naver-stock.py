# 1. 라이브러리 호출
import requests
from bs4 import BeautifulSoup as bs
import openpyxl
import datetime

marketType = {
  "KOSPI" : "0",
  "KOSDAQ" : "1",
}

dt_now = datetime.datetime.now()

# KOSPI + KOSDAQ
for market, code in marketType.items():
  if(market == "KOSPI"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = market
  else:
    wb = openpyxl.load_workbook("C:/Users/welgram-Inwoo/Desktop/네이버_증권_크롤링/"+str(dt_now.date())+".xlsx")
    wb.create_sheet(market)
    ws = wb[market]

  # 컬럼 정의
  ws['A1'] = "순위"
  ws['B1'] = "종목명"
  ws['C1'] = "현재가"
  ws['D1'] = "전일비"
  ws['E1'] = "등락률"
  ws['F1'] = "액면가"
  ws['G1'] = "시가총액"
  ws['H1'] = "자산총계"
  ws['I1'] = "부채총계"
  ws['J1'] = "영업이익"
  ws['K1'] = "PER"
  ws['L1'] = "PBR"
  ws['M1'] = "시가총액/영업이익"
  ws['N1'] = "순자산"
  ws['O1'] = "시가총액/순자산"

  for page in range(1,36):
    # 2. 데이터 호출
    req = requests.get(f"https://finance.naver.com/sise/sise_market_sum.nhn?sosok={code}&page={page}")
    html = req.text
    soup = bs(html, "lxml")

    # 3. 데이터 추출 (파싱) 단계
    stockContents = soup.select("#contentarea > div.box_type_l > table.type_2 > tbody > tr")
    for stockContent in stockContents:
      try:
        objRank = stockContent.select_one("td.no").text                                 # 순위
        objName = stockContent.select_one("td:nth-child(2) >a").text                    # 종목명
        objCurrentPrice = stockContent.select_one("td:nth-child(3)").text               # 현재가
        objFullTime = stockContent.select_one("td:nth-child(4) > span").text            # 전일비
        objFluctuationRate= stockContent.select_one("td:nth-child(5) > span").text      # 등락률
        objFaceValue = stockContent.select_one("td:nth-child(6)").text                  # 액면가
        objCap = stockContent.select_one("td:nth-child(7)").text                        # 시가총액
        objTotalAssets = stockContent.select_one("td:nth-child(8)").text                # 자산총계
        objTotalDebt = stockContent.select_one("td:nth-child(9)").text                  # 부채총계
        objOperatingProfit = stockContent.select_one("td:nth-child(10)").text           # 영업이익
        objPer = stockContent.select_one("td:nth-child(11)").text                       # PER
        objPbr = stockContent.select_one("td:nth-child(12)").text                       # PBR
        #  시가총액 / 영업이익
        # additionalFormulaOne = float(objCap.replace(',','')) / float(objOperatingProfit.replace(',','')) if (objCap != "N/A" and objOperatingProfit != "N/A") else "N/A"
        # #  자산총계 - 부채총계 = 순자산
        # additionalFormulaTwo = float(objTotalAssets.replace(',','')) / float(objTotalDebt.replace(',','')) if (objTotalAssets != "N/A" and objTotalDebt != "N/A") else "N/A"
        # #  시가총액 / 순자산
        # additionalFormulaThree = float(objCap.replace(',','')) / float(additionalFormulaTwo.replace(',','')) if (objCap != "N/A" and additionalFormulaTwo != "N/A") else "N/A"
        # print("additionalFormulaOne :: " + additionalFormulaOne)
        # print("additionalFormulaTwo :: " + additionalFormulaTwo)
        # print("additionalFormulaThree :: " + additionalFormulaThree)

        ws['A'+str(int(objRank)+1)] = objRank
        ws['B'+str(int(objRank)+1)] = objName
        ws['C'+str(int(objRank)+1)] = objCurrentPrice
        ws['D'+str(int(objRank)+1)] = objFullTime
        ws['E'+str(int(objRank)+1)] = objFluctuationRate
        ws['F'+str(int(objRank)+1)] = objFaceValue
        ws['G'+str(int(objRank)+1)] = objCap
        ws['H'+str(int(objRank)+1)] = objTotalAssets
        ws['I'+str(int(objRank)+1)] = objTotalDebt
        ws['J'+str(int(objRank)+1)] = objOperatingProfit
        ws['K'+str(int(objRank)+1)] = objPer
        ws['L'+str(int(objRank)+1)] = objPbr

        print(f"{market} : {objRank}등 : 크롤링중....")
      except AttributeError:
        continue
      except ZeroDivisionError:
        continue
  print(f"{market} : 크롤링완료")

  wb.save("C:/Users/welgram-Inwoo/Desktop/네이버_증권_크롤링/"+str(dt_now.date())+".xlsx")

