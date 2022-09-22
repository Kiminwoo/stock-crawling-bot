from discordwebhook import Discord
from datetime import date


"""
 디스코드에 알림 함수 
  Args : 
    rankKospiList (objecList) : 코스피 상품 객체 리스트
    rankKosdaqList (objectList) : 코스닥 상품 객체 리스트 
 
"""
def sendDiscord(rankKospiList,rankKosdaqList):

    discord = Discord(url="")

    discord.post(
        embeds= [
            {
                "author" :{
                    "name" : "naver-stock-bot" , 
                    "icon_url" : "https://picsum.photos/400/300"
                }  , 
                "title" : "오늘의 네이버 주식 현황 [ "+str(dt_now.date())+" ]" , 
                "color": 0,
            } , 


            { 
              "title" : "코스피 [ 1등 ] : " +rankKospiList[0].name , 
              "color" : 1127128,
              "fields": [
                    {"name": "현재가", "value": rankKospiList[0].currentPrice, "inline": True},
                    {"name": "전일비", "value": rankKospiList[0].fullTime, "inline": True},
                    {"name": "등락률", "value": rankKospiList[0].flucTuationRate, "inline": True},

                    {"name": "액면가", "value": rankKospiList[0].faceValue, "inline": True},
                    {"name": "시가총액", "value": rankKospiList[0].cap, "inline": True},
                    {"name": "자산총계", "value": rankKospiList[0].totalAssets, "inline": True},
                    {"name": "부채총계", "value": rankKospiList[0].totalDebt, "inline": True},
                    {"name": "영업이익", "value": rankKospiList[0].operatingProfit, "inline": True},
                    {"name": "PER", "value": rankKospiList[0].per, "inline": True},
                    {"name": "PBR", "value": rankKospiList[0].pbr, "inline": True},
                ]
            } , 


            { 
              "title" : "코스피 [ 2등 ] : " +rankKospiList[1].name , 
              "color" : 1127128,
              "fields": [
                    {"name": "현재가", "value": rankKospiList[1].currentPrice, "inline": True},
                    {"name": "전일비", "value": rankKospiList[1].fullTime, "inline": True},
                    {"name": "등락률", "value": rankKospiList[1].flucTuationRate, "inline": True},

                    {"name": "액면가", "value": rankKospiList[1].faceValue, "inline": True},
                    {"name": "시가총액", "value": rankKospiList[1].cap, "inline": True},
                    {"name": "자산총계", "value": rankKospiList[1].totalAssets, "inline": True},
                    {"name": "부채총계", "value": rankKospiList[1].totalDebt, "inline": True},
                    {"name": "영업이익", "value": rankKospiList[1].operatingProfit, "inline": True},
                    {"name": "PER", "value": rankKospiList[1].per, "inline": True},
                    {"name": "PBR", "value": rankKospiList[1].pbr, "inline": True},
                ]
            } , 


            {
                "title" : "코스피 [ 3등 ] : " +rankKospiList[2].name , 
                "color" : 1127128,
                "fields": [
                    {"name": "현재가", "value": rankKospiList[2].currentPrice, "inline": True},
                    {"name": "전일비", "value": rankKospiList[2].fullTime, "inline": True},
                    {"name": "등락률", "value": rankKospiList[2].flucTuationRate, "inline": True},

                    {"name": "액면가", "value": rankKospiList[2].faceValue, "inline": True},
                    {"name": "시가총액", "value": rankKospiList[2].cap, "inline": True},
                    {"name": "자산총계", "value": rankKospiList[2].totalAssets, "inline": True},
                    {"name": "부채총계", "value": rankKospiList[2].totalDebt, "inline": True},
                    {"name": "영업이익", "value": rankKospiList[2].operatingProfit, "inline": True},
                    {"name": "PER", "value": rankKospiList[2].per, "inline": True},
                    {"name": "PBR", "value": rankKospiList[2].pbr, "inline": True},
                ]
            } , 


            { 
              "title" : "코스닥 [ 1등 ] : " +rankKosdaqList[0].name , 
              "color" : 14177041,
              "fields": [
                    {"name": "현재가", "value": rankKosdaqList[0].currentPrice, "inline": True},
                    {"name": "전일비", "value": rankKosdaqList[0].fullTime, "inline": True},
                    {"name": "등락률", "value": rankKosdaqList[0].flucTuationRate, "inline": True},

                    {"name": "액면가", "value": rankKosdaqList[0].faceValue, "inline": True},
                    {"name": "시가총액", "value": rankKosdaqList[0].cap, "inline": True},
                    {"name": "자산총계", "value": rankKosdaqList[0].totalAssets, "inline": True},
                    {"name": "부채총계", "value": rankKospiList[0].totalDebt, "inline": True},
                    {"name": "영업이익", "value": rankKosdaqList[0].operatingProfit, "inline": True},
                    {"name": "PER", "value": rankKosdaqList[0].per, "inline": True},
                    {"name": "PBR", "value": rankKosdaqList[0].pbr, "inline": True},
                ]
            } , 


            { 
              "title" : "코스닥 [ 2등 ] : " +rankKosdaqList[1].name , 
              "color" : 14177041,
              "fields": [
                    {"name": "현재가", "value": rankKosdaqList[1].currentPrice, "inline": True},
                    {"name": "전일비", "value": rankKosdaqList[1].fullTime, "inline": True},
                    {"name": "등락률", "value": rankKosdaqList[1].flucTuationRate, "inline": True},

                    {"name": "액면가", "value": rankKosdaqList[1].faceValue, "inline": True},
                    {"name": "시가총액", "value": rankKosdaqList[1].cap, "inline": True},
                    {"name": "자산총계", "value": rankKosdaqList[1].totalAssets, "inline": True},
                    {"name": "부채총계", "value": rankKosdaqList[1].totalDebt, "inline": True},
                    {"name": "영업이익", "value": rankKosdaqList[1].operatingProfit, "inline": True},
                    {"name": "PER", "value": rankKosdaqList[1].per, "inline": True},
                    {"name": "PBR", "value": rankKosdaqList[1].pbr, "inline": True},
                ]
            } , 


            {
                "title" : "코스닥 [ 3등 ] : " +rankKosdaqList[2].name , 
                "color" : 14177041,
                "fields": [
                    {"name": "현재가", "value": rankKosdaqList[2].currentPrice, "inline": True},
                    {"name": "전일비", "value": rankKosdaqList[2].fullTime, "inline": True},
                    {"name": "등락률", "value": rankKosdaqList[2].flucTuationRate, "inline": True},

                    {"name": "액면가", "value": rankKosdaqList[2].faceValue, "inline": True},
                    {"name": "시가총액", "value": rankKosdaqList[2].cap, "inline": True},
                    {"name": "자산총계", "value": rankKosdaqList[2].totalAssets, "inline": True},
                    {"name": "부채총계", "value": rankKosdaqList[2].totalDebt, "inline": True},
                    {"name": "영업이익", "value": rankKosdaqList[2].operatingProfit, "inline": True},
                    {"name": "PER", "value": rankKosdaqList[2].per, "inline": True},
                    {"name": "PBR", "value": rankKosdaqList[2].pbr, "inline": True},
                ]
            } 
        ],
    )

    discord.post(
        file={
            "file1" : open("C:/Users/welgram-Inwoo/Desktop/네이버_증권_크롤링/"+str(dt_now.date())+".xlsx","rb"),
        }
    )