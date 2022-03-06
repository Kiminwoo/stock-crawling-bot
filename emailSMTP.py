import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

naver_id =''					#네이버 아이디 입력
naver_pass = ''			#패스워드 입력

def makeFrom(sender,receiver,title,content,filename):
  msg = MIMEMultipart('alternative')
  msg['Subject'] = "%s"%(title)
  msg['From'] = sender
  msg['To'] = receiver

  # 메일 내용 쓰기
  part2 = MIMEText(content, 'plain')
  msg.attach(part2)

  # 첨부파일 추가
  part = MIMEBase('application', "octet-stream")
  with open("C:/Users/welgram-Inwoo/Desktop/네이버_증권_크롤링/"+filename, 'rb') as file:
    part.set_payload(file.read())
  encoders.encode_base64(part)
  part.add_header('Content-Disposition', "attachment", filename=filename+".xlsx")
  msg.attach(part)

  return msg.as_string()

def send(receiver, title, filename):
  # 로그인하기
  server = smtplib.SMTP_SSL('smtp.naver.com',465)
  server.login(naver_id,naver_pass)

  body = makeFrom(naver_id,receiver,title,"",filename)

  # 메일 보내고 서버 끄기
  server.sendmail(naver_id,receiver,body)
  server.quit()

  return True

if __name__ == '__main__':
  reveiver_email = 'dlsdn166@gmail.com'
  title = "email 테스트";
  email_result = send(reveiver_email,title)
  print(email_result)
