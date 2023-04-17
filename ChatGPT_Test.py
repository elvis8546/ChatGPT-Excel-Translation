import os
import openai
import xlwings as xw

wb = xw.Book('') #The Excel file path
wks = xw.sheets  #To get Excel sheets' line
ws = wks[0]
ws2 = wks[1]

openai.api_key = f''  #The ChatGPT API key (premium version required)
for i in range(1528):
    w = i+1289
    column =  "U"+str(w)
    column2 = "A"+str(w)
    content = ws.range(column).value
    print("A value in sheet1 :", content)   
    completion = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",
    messages=[
            {"role": "system", "content": "系統訊息，目前無用"},
            {"role": "assistant", "content": "此處填入機器人訊息"},
            {"role": "user", "content": content}
        ]
    )
    print(completion.choices[0].message.content)
    ans = completion.choices[0].message.content
    ws2[column2].value = ans
    print("第"+column2+"欄翻譯完成")
        