import os
import sys
if not sys.warnoptions:
    import warnings
    warnings.simplefilter("ignore")
from io import BytesIO
import pandas as pd
from datetime import datetime
import xml.etree.ElementTree as ET
from flask import Flask, render_template, request,send_file
from werkzeug.utils import secure_filename
app = Flask(__name__)
##x=pd.DataFrame()
@app.route('/')
def upload_file():
   return render_template('index.html')
@app.route('/upload', methods = ['GET', 'POST'])
def uploadfile():
   if request.method == 'POST':
      f = request.files['file']
      tree = ET.parse(f)
      root = tree.getroot()
      List = []
      List1 = []
      for i in root.findall('.//TALLYMESSAGE/*'):
         dict = {}
         for j in i:
            if (len(j.getchildren()) > 0):
               for k in j.getchildren():
                  dict.update({j.tag + k.tag: k.text})
                  if (len(k.getchildren()) > 0):
                     for z in k.getchildren():
                        dict.update({k.tag + z.tag: z.text})
                  else:
                     dict.update({j.tag + k.tag: k.text})
            else:
               dict.update({j.tag: j.text})
         List.append(dict)

      df1 = pd.DataFrame()
      for i in range(len(List)):
         df1 = df1.append(List[i], ignore_index=True)
      df2 = pd.DataFrame()
      for i in df1.index:
         dict2 = {'Date': df1['DATE'][i], 'TransactionType': '',
              'VchNo': df1['VOUCHERNUMBER'][i], 'RefNo': df1.get('REFERENCE')[i],
              'RefType': df1.get('BILLALLOCATIONS.LISTBILLTYPE')[i], 'RefDate': df1.get('REFERENCEDATE')[i],
              'Debtor': df1.get('PARTYNAME')[i], 'RefAmount': df1.get('BILLALLOCATIONS.LISTAMOUNT')[i],
              'Amount': df1.get('LEDGERENTRIES.LISTAMOUNT')[i], 'Particulars': df1.get('PARTYLEDGERNAME')[i],
              'VchType': df1.get('VOUCHERTYPENAME')[i],
              'AmountVerified': df1.get('LEDGERENTRIES.LISTISPARTYLEDGER')[i]}
         df2 = df2.append(dict2, ignore_index=True)
      list3 = list(df2['Particulars'].dropna().unique())
      v = []
      for i in df2.index:
         if (df2['Particulars'][i] not in list3):
            df2['TransactionType'][i] = 'Others'
         elif (df2['Particulars'][i] not in v):
            df2['TransactionType'][i] = 'Parent'
            v.append(df2['Particulars'][i])
         else:
            df2['TransactionType'][i] = 'child'
    
      df2['Date'] = pd.to_datetime(df2['Date'], format='%Y%m%d')
    
      
      output= BytesIO()
      writer = pd.ExcelWriter(output, engine='xlsxwriter')
      df2.to_excel(writer, sheet_name='Sheet1')
      writer.save()
      output.seek(0)
      xlsx_data = output.getvalue()
      return send_file(output, attachment_filename='output.xlsx', as_attachment=True) 
     
if __name__ == '__main__':
   app.run()
