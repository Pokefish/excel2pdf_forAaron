# pip install xlsxwriter,requests,pandas,pywin32
import os
import requests
import pandas as pd
import xlsxwriter
from win32com.client import DispatchEx

# https://www.ragic.com/intl/en/doc-api
def getRagic():
     url = ''
     headers = {
     'Authorization': '', 
     'Authorization': '',
     "content-type" : "application/json",
     "charset":"utf-8"
     }
     response = requests.request(
     'GET',url,
     headers=headers
     )
     re = response.json()
     
     return re

def chooseData(date,department,entity,caseType):
     test = pd.DataFrame(getRagic())
     # 拿掉後面的（人壽）
     for i in test:
          a = test[i]['案件類型'].split('(',1)
          test[i]['案件類型'] = a[0]

     test = test.T 
     test = test.loc[test['受理日期'] >= "2021/06/01" ]  #新制
     if date == [''] or date == '': #如果參數沒有指定某一個的話，當作全選
          test = test.loc[test['受理日期'] != '']  
     else:
          test = test.loc[test['受理日期'] == date]          #篩選日期 date
     if department == [''] or department == '':
          test = test.loc[test['所屬單位'] != '']
     else:
          test = test.loc[test['所屬單位'] == department]    #單位 ,department
     if entity == [''] or entity == '':
          test = test.loc[test['保險公司'] != ''] 
     else:
          test = test.loc[test['保險公司'] == entity ]        #篩選公司 ,entity
     

     if caseType == [''] or caseType == '':
          test = test.loc[test['案件類型'] != '']      
     else:
          test = test.loc[test['案件類型'] == caseType]      #案件類型 ,caseType
     
     test = test.T
     
     print("資料筆數：",len(test.columns),f"\n日期：{date}\n單位：{department}\n人壽：{entity}\n案件類型：{caseType}")
     return test 



def mixtype(test,entity,caseType):

     test = test.drop(['_ragicId',"_star","_index_title_","要被保人同一人","_index_calDates_",
          '_subtable_1000255','送件時間','受理狀態','上傳檔案','資料管理者','_index_','_seq',
          '未命名'],axis=0) # 本來就用不到的
     
     test2 = test.drop(['附件一','附件二','附件三','附件四','附件五','附件六','附件七','附件八','附件九','附件十',
                    '所屬單位', '保險公司', '案件類型', '受理日期','退件因素'],axis = 0) 

     d1 = test.loc["附件一":"附件十"] # 用這個才不會跑掉
     d1 = d1.to_dict('dict')
     id = list(d1.keys()) # id

     sub = []
     for i in range(len(id)):
          a = d1[id[i]].values()
          sub_ls = list(a for a in a if a[:] != '') # sub 有 ''  #現在只能拿到最後一個
          sub.append(sub_ls)
     

     if '人壽' in  entity :
          if '遠雄人壽' in entity or '台灣人壽' in entity or '全球人壽' in entity :
               csv = entity +'.csv'
          else :
               csv = '其餘壽險公司.csv'
     elif '產物' in entity:
          csv  =   '所有產險公司.csv'
     
     
     data = pd.read_csv("C:/Users/user/Desktop/forAaron/mapping/"+csv) #將附件表丟進來
     
     far = data[caseType].dropna(axis= 'index')  # 拿自己需要的type  #還要砍掉nan
     
     test3 = pd.concat([test2,pd.DataFrame(index=far)]) # 屬性資料跟附件合體

     for i in range(len(id)) :
          test3.at[sub[i],id[i]]='V' # a_row \ b_col #如果真實資料有誤，多了附件但mapping表還沒更新的話，就會有ERROR # 要用try 去試
          
     test4 = test3.T
     num_ls = []

     for i in range(0,len(test4)):
          num = i+1 
          num_ls.append(num)
     test4.pop('其他')
     memo = test4.pop('備註欄')
     test4.insert(test4.shape[1],memo.name,memo)
     test4 = test4.reset_index(drop=True)
     test4.index.name = '序號'
     
     return test4

def xslmformat(test4,date,department,entity,caseType,title):

     df = test4
     d = date[0:4]+date[5:7]+date[8:]
     #單位代號表（加在表頭）產物公司都沒有
     departno = pd.read_csv("C:/Users/user/Desktop/forAaron/mapping/單位代號.csv") 
     departno = pd.DataFrame(departno)
     departno.set_index("保險公司" , inplace=True)
     try :
          no = departno[department][entity]
     except KeyError:
          no = ''

     #產險壽險處理
     if '產物' in entity:
          entity_short = entity[0:2]+'產'
     else :
          entity_short = entity[0:2]
     
     # info/表頭/名稱
     if department == 'XXX1':
          info = '單位：XXX1 '+no+'\n電話：XXX1\n傳真：XXX1\n助理受理日期：'+date 
          name = d+'XXX1'+title+'-'+entity_short+'.xlsx'
     
     elif department == 'XXX2':
          info = '單位：XXX2 '+no+'\n電話：XXX2\n傳真：XXX2\n助理受理日期：'+date 
          name = d+'XXX2'+title+'-'+entity_short+'.xlsx'
     else:
          info = '單位：XXX3 '+no+'\n電話：XXX3\n傳真：XXX3\n助理受理日期：'+date
          name = d+'XXX3'+title+'-'+entity_short+'.xlsx'
     
     sheetname = caseType

     writer = pd.ExcelWriter('C:\\Users\\user\\Desktop\\forAaron\\printing\\'+name)
     df.to_excel(writer, sheet_name=sheetname, index=True, na_rep=' ')    

     workbook  = writer.book
     worksheet = writer.sheets[sheetname]

     ## 第一列 ＃關掉headerF不會讓綠綠也影響大家 
     header_format = workbook.add_format({
          'text_wrap': True, #自動換行
          'align': 'center', #文字置中
          'valign': 'vcenter', #置中對齊
          'top': 2, #上粗黑框
          'bottom': 2, #下粗黑框
          'font_size':12,
          'left':1,
          'right':1,
          'bold':False
          }) 
     sub_format = workbook.add_format({
          'text_wrap': True, #自動換行
          'align': 'center', #文字置中
          'valign': 'vcenter', #置中對齊
          'top': 2, #上粗黑框
          'bottom': 2, #下粗黑框
          'left':1,
          'right':1,
          'font_size':9,
          })
     #格子###
     cell_format = workbook.add_format({
          'font_size':12,
          'align': 'center', #文字置中
          'valign': 'vcenter', #置中對齊
          'border':1,
          'bold':False
          })

     for column in df:
          sub_uni = max(df.columns.astype(str).map(len).max(),
                    len(df[column])) #附件字最多12 ##
          worksheet.set_row(0,(sub_uni+1)*12) #第一row 依據附件字數max做為高度


     worksheet.set_column(1,len(df.columns),18,cell_format) ##每一col都先設定，寬度18，字體大小12（後面調整「附件寬度」）
     worksheet.set_column(5,len(df.columns)-1,3,cell_format)  ## 附件col，寬度3。


     ##為了讓序號列都兩邊粗框
     in0_format = workbook.add_format({
          'text_wrap': True, #自動換行
          'align': 'center', #文字置中
          'valign': 'vcenter', #置中對齊
          'font_size':12,
          'left':2,
          'right':2,
          'top': 2, #上粗黑框
          'bottom': 2, #下粗黑框
          'bold':False
          }) 
     in_format = workbook.add_format({
          'text_wrap': True, #自動換行
          'align': 'center', #文字置中
          'valign': 'vcenter', #置中對齊
          'font_size':12,
          'border':1,
          'right':2,
          'left':2,
          'bold':False
          }) 

     #最後「備註」的右邊粗框
     la0_format = workbook.add_format({
          'text_wrap': True, #自動換行
          'align': 'center', #文字置中
          'valign': 'vcenter', #置中對齊
          'border':1,
          'right':2,
          'bottom':2,
          'top':2,
          'bold':False
          }) 
     la_format = workbook.add_format({
          'text_wrap': True, #自動換行
          'align': 'center', #文字置中
          'valign': 'vcenter', #置中對齊
          'border':1,
          'right':2,
          'bold':False
          }) 

     larow_format =workbook.add_format({
          'font_size':12,
          'align': 'center', #文字置中
          'valign': 'vcenter', #置中對齊
          'border':1,
          'bottom':2,
          'bold':False
          })
     for row_num in range(len(df)) : 
          worksheet.set_row(row_num+1,24) # 內容高度24
          worksheet.write(row_num+1,0,str(row_num+1),in_format) #cell_format) #序號數字

     # col_num普通走訪Ａ～最後一列
     for col_num, value in enumerate(df.columns.values):
          if col_num+1 < 5 or col_num >= len(df.columns)-1:
               worksheet.set_column(1,len(df.columns),18,cell_format) ## 欄位寬度 ##
               worksheet.set_column(0,0,3) ## 序號寬度 ## 有框線但字體沒變
               worksheet.write(0, col_num+1, value, header_format) #+1因為要留index，不要蓋掉 
          else:
               worksheet.write(0, col_num+1, value, sub_format)  # 附件內容


     # 特殊粗框製造
     worksheet.write(0, 0, '序號',in0_format)  ## 序號本人
     worksheet.write(0, len(df.columns), df.columns[-1], la0_format) ## 備註本人
     worksheet.set_column(len(df.columns),len(df.columns),18,la_format) ## 備註的dada
     worksheet.set_row(len(df)+1,1,larow_format) ##最底下

     ## 測試列印設定 #########

     worksheet.set_paper(9) #A4大小
     worksheet.set_landscape() #橫向
     worksheet.fit_to_pages(1, 0) # 縮成一頁寬
     # 列印格範圍 df所有欄位、df所有資料集
     worksheet.set_h_pagebreaks([len(df)+2]) 
     worksheet.set_v_pagebreaks([len(df.columns)])
     # 設定邊界 #單位inches
     worksheet.set_margins(left=0.24, right=0.24, top=1.2, bottom=0.75) 
     # worksheet.set_header() 表頭的東西 /XX人壽 / ＸＸＸＸ書  / 資訊欄 
     where = entity
     # title = '新契約照會回覆文件清單' #test['未命名']
     
     title = title
     

     

     worksheet.set_header('&L&36&B&"Microsoft JhengHei"'+where+'&C&16&B&"Microsoft JhengHei"'+title+'&R&10&"Microsoft JhengHei"'+info)
     #頁尾的東西#/ / /第Ｎ頁、共Ｍ頁
     worksheet.set_footer('&R&12&"Microsoft JhengHei" 第&P頁/共&N頁') 

     ##待辦####


     writer.save()


def excel_pdf():

      path = r'C:\\Users\\user\\Desktop\\forAaron\\' 

      # 列出文件夹里面所有文件

      list_path = os.listdir(path)
      # 过滤得到xlsx文件

      wait_turn_list  =[i for i in list_path if 'xlsx' == i[-4:] and  '~$' != i[:2]]
      for i in wait_turn_list:
           #准确的文件路径并转换xlsx-->pdf
           pdf_path = path+i.replace('xlsx', 'pdf')
      #对不同的格式进行设定（xlsx,word等）
           xlApp = DispatchEx("Excel.Application")
      #对表格进行操作
           books = xlApp.Workbooks.Open(path+i)
           books.ExportAsFixedFormat(0, pdf_path)
           xlApp.Quit()
      return 'over'



if __name__ == "__main__":
     main()
