import os
from ragic import chooseData,mixtype,xslmformat
from flask import Flask,render_template,request
import datetime
from win32com.client import DispatchEx
import pythoncom
app = Flask(__name__)

# filter()\submit()都要存在才可以打開
#filter 路由 
@app.route("/", methods=['GET','POST']) #放func的路徑 ，可以多個
def select():
     
     if request.method == 'POST':
          starttime = datetime.datetime.now()
          date1 = request.form.getlist('Date')
          # revert to yyyy/mm/dd
          date = []
          for all in range(len(date1)):
               a = datetime.datetime.strptime(date1[all], '%Y-%m-%d').strftime('%Y/%m/%d')
               date.append(a)
          
          department = request.form.getlist('Depart')
          entity = request.form.getlist('Entity')
          caseType = request.form.getlist('caseType')
          
          # 主程式
          submit(date,department,entity,caseType)
          excel_pdf()
          # 計算時長
          endtime = datetime.datetime.now()
          long = (endtime - starttime).seconds
          minute = long//60
          sec = long%60

          return render_template('submit.html',**locals()) 
     return render_template('filter.html') 
     #這樣這邊就可以retun html 但要記得import render_template

def excel_pdf():

      path = r'C:\\Users\\user\\Desktop\\forAaron\\printing\\' 

      # 列出文件夹里面所有文件

      list_path = os.listdir(path)
      # 过滤得到xlsx文件

      wait_turn_list  =[i for i in list_path if 'xlsx' == i[-4:] and  '~$' != i[:2]]
      for i in wait_turn_list:
           #准确的文件路径并转换xlsx-->pdf
           pdf_path = path+i.replace('xlsx', 'pdf')
      #对不同的格式进行设定（xlsx,word等）
           pythoncom.CoInitialize()#加上的 #pywintypes.com_error: (-2147221008, 'CoInitialize 尚未被呼叫。', None, None)
           xlApp = DispatchEx("Excel.Application")
           pythoncom.CoInitialize()#加上的

      #对表格进行操作
           books = xlApp.Workbooks.Open(path+i)
           books.ExportAsFixedFormat(0, pdf_path)
           xlApp.Quit()
      return 'over'
     
# submit的路由
@app.route("/submit", methods=['GET','POST']) #"必定得 /xxxxx 不然無法跳"
def submit(date,department,entity,caseType):
     # date yyyy-mm-dd要轉成yyyy/mm/dd
     if department == ['全選'] :
          department = ['XXX1','XXX2','XXX3']
     if entity == ['全選']:
          entity = ["甲","乙","丙"]
     elif entity == ['其餘壽險公司']:
          entity = ["甲Ａ","甲Ｂ","甲Ｃ"]     
     elif entity == ['所有產險公司']:
          entity = ["A1","A2","A3"]
     
     if caseType == ['全選'] :
          caseType=["A","B","C"]
          

     for i in range(len(date)):
          for j in range(len(department)):
               for n in range(len(entity)):
                    for m in range(len(caseType)):
                         test = chooseData(date[i],department[j],entity[n],caseType[m])
                         if len(test.columns) != 0: 
                              test4 = mixtype(test,entity[n],caseType[m])
                              title = test.T['未命名'][0]
                              xslmformat(test4,date[i],department[j],entity[n],caseType[m],title)

     
     # return render_template('submit.html',**locals()) 

#####RAGIC#######

if __name__ == '__main__':
     app.debug = True 
     app.run()
     # main()






