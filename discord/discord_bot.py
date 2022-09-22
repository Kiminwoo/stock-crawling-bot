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
                "fields": [
                    {"name": "kospi 1등", "value": rankKospiList[0].name, "inline": True},
                    {"name": "kospi 2등", "value": rankKospiList[1].name, "inline": True},
                    {"name": "kospi 3등", "value": rankKospiList[2].name, "inline": True},

                    {"name": "kosdaq 1등", "value": rankKosdaqList[0].name, "inline": True},
                    {"name": "kosdaq 2등", "value": rankKosdaqList[1].name, "inline": True},
                    {"name": "kosdaq 3등", "value": rankKosdaqList[2].name, "inline": True},

                ],
            }
        ],
    )

    discord.post(
        file={
            "file1" : open("C:/Users/welgram-Inwoo/Desktop/네이버_증권_크롤링/"+str(dt_now.date())+".xlsx","rb"),
        }
    )