# 1. 라이브러리 호출
import requests
from bs4 import BeautifulSoup as bs
import openpyxl
import datetime
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import time

from emailSMTP import send


dt_now = datetime.datetime.now()

selectList = ["거래량","시가총액(억)","PER(배)","자산총계(억)","부채총계(억)","PBR(배)"]

# 페이징 스킵 함수
def skipList(content):
  skipContentList = ["맨앞","다음","맨뒤","이전"]

  if content in skipContentList:
    return False
  else:
    return True

# 페이징 이동 함수
def pagingMove(page,driver,cssSelect):
  for pageListCount in cssSelect:
    if(page in (11,21,31)):
      driver.find_element_by_css_selector("#contentarea > div.box_type_l > table.Nnavi > tbody > tr > td.pgR > a").click()
      break
    else:
      if(skipList(pageListCount.text.replace(" ",""))):
        if(int(pageListCount.text) == page):
          pageListCount.click()
          time.sleep(2)
          break

# topList 초기화
def topListReset(cssSelect):
  for list in cssSelect:
    tdList = list.find_elements_by_tag_name("td")
  for inputTag in tdList:
    try:
      if(inputTag.find_element_by_tag_name("input").get_attribute('checked')):
        inputTag.find_element_by_tag_name("input").click()
    except NoSuchElementException:
      print("No such")

# 검색조건 변경
def search(cssSelect):
  for list in cssSelect:
    tdList = list.find_elements_by_tag_name("td")
  for inputTag in tdList:
    try:
      if(inputTag.text in selectList):
        inputTag.find_element_by_tag_name("input").click()
    except NoSuchElementException:
      print("No such")

def crawlBot():

  marketType = {
    "KOSPI" : "0",
    "KOSDAQ" : "1",
  }

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

    # 변동적인 칼럼값 정의
    objRankCount = 1

    # 드라이버 연결
    driver = webdriver.Chrome()

    # 웹사이트 이동
    driver.get(f"https://finance.naver.com/sise/sise_market_sum.nhn?sosok={code}&page=1")

    topList = driver.find_elements_by_css_selector("#contentarea_left > div.box_type_m > form > div > div > table > tbody > tr")

    # topList 초기화
    topListReset(topList)

    # 검색조건 변경
    search(topList)

    # 검색조건 서치
    driver.find_element_by_css_selector("#contentarea_left > div.box_type_m > form > div > div > div > a:nth-child(1) > img").click()
    time.sleep(3)

    for page in range(1,36):

      # 페이징 이동
      pageList = driver.find_elements_by_css_selector("#contentarea > div.box_type_l > table.Nnavi > tbody > tr > td")
      pagingMove(page,driver,pageList)

      # 2. 데이터 호출
      req = driver.page_source
      soup = bs(req, "lxml")

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

          # if(objPer == "N/A" or objPbr == "N/A"): # PER or PBR N/A일 경우 해당 종목 스킵
          #   objRankCount == 0
          #   continue
          # else:
          #   objRankCount == 1

          ws['A'+str(int(objRank)+objRankCount)] = objRank
          ws['B'+str(int(objRank)+objRankCount)] = objName
          ws['C'+str(int(objRank)+objRankCount)] = objCurrentPrice
          ws['D'+str(int(objRank)+objRankCount)] = objFullTime
          ws['E'+str(int(objRank)+objRankCount)] = objFluctuationRate
          ws['F'+str(int(objRank)+objRankCount)] = objFaceValue
          ws['G'+str(int(objRank)+objRankCount)] = objCap
          ws['H'+str(int(objRank)+objRankCount)] = objTotalAssets
          ws['I'+str(int(objRank)+objRankCount)] = objTotalDebt
          ws['J'+str(int(objRank)+objRankCount)] = objOperatingProfit
          ws['K'+str(int(objRank)+objRankCount)] = objPer
          ws['L'+str(int(objRank)+objRankCount)] = objPbr

          print(f"{market} : {objRank}등 : 크롤링중....")
        except AttributeError:
          continue
        except ZeroDivisionError:
          continue
    print(f"{market} : 크롤링완료")

    wb.save("C:/Users/welgram-Inwoo/Desktop/네이버_증권_크롤링/"+str(dt_now.date())+".xlsx")
    driver.close()

  return True

if __name__ == '__main__':
  reveiver_email = 'leekh916@hanmail.net'
  title = "[알림] 네이버 주식 크롤링 봇 정보 수집 완료 test [" + str(dt_now.date()) +"]"
  if(crawlBot()):
    try:
      email_result = send(reveiver_email,title,str(dt_now.date()))
      print("이메일발송 : " + str(email_result))
    except Exception as e:
      print(e)