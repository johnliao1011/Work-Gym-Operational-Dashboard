import win32com.client
import os
import numpy as np
import pandas as pd
from pretty_html_table import build_table
from pathlib import Path

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts

# search for the target folders
# specify the account in outlook (ex."IMAP_john.liao@cathaylife.com.tw")
folders = outlook.Folders["IMAP_john.liao@cathaylife.com.tw"].Folders["收件匣"].Folders["空音警示信"]


# store the content of mail in the folder
mail_list = []
for i in folders.Items:
    mail_list.append(i.Body)


## create new dictionary to store the message        
dict = {"Alarm_No":[],"Alarm_Service_Code":[], "Alarm_Description":[], "Alarm_Time":[]}

## store the message in the dictionary
for i in range(len(mail_list)):
    
    message = mail_list[i]
    info_break = message.split("Alarm ")

    content_word = []


    info_break2 = [infor.strip().split(": ") for infor in info_break]
    for content in info_break2:
        for subcontent in content:
            content_word.append(subcontent)

    dict["Alarm_No"].append(content_word[content_word.index("No.")+1].strip())
    dict["Alarm_Service_Code"].append(content_word[content_word.index("Service Code")+1].strip())
    try:
        phone_number = content_word[content_word.index("Description")+1].strip().split(" ")[0].split(":")
        dict["Alarm_Description"].append(phone_number[phone_number.index("ExtNo")+1].strip())
    except ValueError:
        dict["Alarm_Description"].append(np.NaN)
    dict["Alarm_Time"].append(content_word[len(content_word)-1].split("\r")[0].split("Time:")[1])

# convert latest message into dataframe and change the type
new_df = pd.DataFrame(dict)
new_df["Alarm_No"] = pd.to_numeric(new_df["Alarm_No"])
new_df["Alarm_Description"] = pd.to_numeric(new_df["Alarm_Description"])
new_df["Alarm_Time"] = pd.to_datetime(new_df["Alarm_Time"])
new_df.insert(0, 'Date', new_df["Alarm_Time"].dt.date)

result = new_df[["Date", "Alarm_Description", "Alarm_Time"]].groupby(["Date", "Alarm_Description"], as_index=False).count().sort_values(by=['Date', "Alarm_Time"], ascending=False)


# merge statistic result with person and group
group = pd.read_excel("話機使用者.xlsx")
group = group.rename(columns = {"實體分機":"Alarm_Description"})
final_result = pd.merge(result, group, on="Alarm_Description", how = "left")
final_result["Alarm_Description"] = final_result["Alarm_Description"].astype(pd.Int64Dtype())
final_result = final_result.rename(columns= {"Date":"發生日期", "Alarm_Description":"話機", "Alarm_Time":"次數"})


# merge list with person and group
history_list = pd.read_excel("歷史清單.xlsx")
final_list = pd.merge(new_df, group, on="Alarm_Description", how = "left")
final_list["Alarm_Description"] = final_list["Alarm_Description"].astype(pd.Int64Dtype())

history_list = pd.concat([history_list,final_list])
history_list.to_excel("歷史清單.xlsx", index = False)


final_new_df = pd.merge(new_df, group, on="Alarm_Description", how = "left")
final_new_df = final_new_df.drop(["Date",'Alarm_No', 'Alarm_Service_Code'], axis = 1)

final_new_df = final_new_df[['Alarm_Time', 'Alarm_Description', '客服專員', '單位', '地點', '組別']]
final_new_df['Alarm_Description'] = final_new_df['Alarm_Description'].astype(pd.Int64Dtype()) 
final_new_df = final_new_df.rename(columns= {"Alarm_Description":"空音話機", 'Alarm_Time':'空音發生日期Alarm_Time'})

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)

mail.To = "veronica09@cathaylife.com.tw ; cady0730@cathaylife.com.tw ; gh20502@cathaylife.com.tw ; iba10008@cathaylife.com.tw; chiayu0616@cathaylife.com.tw ; ypqoo@cathaylife.com.tw ; wanting@cathaylife.com.tw ; ho-pingfan@cathaylife.com.tw ; vivian0228@cathaylife.com.tw ; sammy@cathaylife.com.tw ; tabac01@cathaylife.com.tw ; mgkmr12@cathaylife.com.tw ; leon0126@cathaylife.com.tw ; MAYFONG@cathaylife.com.tw ; sugartang@cathaylife.com.tw ; mikowu@cathaylife.com.tw"
mail.CC = "john.liao@cathaylife.com.tw ; tsenying@cathaylife.com.tw ; newpei@cathaylife.com.tw ; su-i-wen@cathaylife.com.tw ; Wells@cathaylife.com.tw ; szuling@cathaylife.com.tw"
mail.Subject = "空音警示信統計"
mail.Subject = "空音警示信統計"

mail.HTMLBody = """
<html>
  <head></head>
  <body>
  	<font color="Dark" size=+1 face="Arial">
    <p>Dear All:<br>
       空音警示信統計如下，若有任何問題請聯繫我，感謝~~<br>
       <br>
       空音警示信寄送規則:<br>
       - 主要錄音(ex.TP-IP)空音4分鐘寄信<br>
       - 備援錄音(ex.TP-IPBK)空音1.5分鐘寄信<br>
       <br>
       統計資訊如下:
       {0:}
       <br>
       <br>
       詳細清單如下:
       {1:}
    </p>
    </font>
  </body>
</html>
""".format((final_result.to_html(index = False, border = 5000, classes="table table-striped")),
           (final_new_df.to_html(index = False, border = 5000, classes="table table-striped")))

mail.BodyFormat = 3
mail.GetInspector 

# mail.Display()
# send mail

mail.Send()
