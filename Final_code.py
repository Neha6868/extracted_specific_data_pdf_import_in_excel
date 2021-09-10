import os
import re
import json
import glob
import openpyxl
import pdfplumber
import pandas as pd
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP

BASE_DIR = Path(__file__).resolve().parent

for filename in glob.glob(os.path.join(BASE_DIR, '*.pdf')):
    if filename.endswith('.pdf'):
        
        fullpath = os.path.join(BASE_DIR, filename)
        
        all_text = ""
        
        with pdfplumber.open(fullpath) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                all_text += '\n' + text
        
        all_lines = list(filter(bool,all_text.split('\n')))
        
        """
        This is for VMLD
        """
        
        dict1 = {}

        list_dec = [i for i,val in enumerate(all_lines) if 'Declaração:' in val]
        di = all_lines[list_dec[0]].split(" ")[1].split("/")[1].split("-")[0]
        dict1['DI'] = [int(di)]

        list_brazil = [i for i,val in enumerate(all_lines) if 'SECRETARIA DA RECEITA FEDERAL DO BRASIL - RFB' in val]
        dict1['Modal'] = [all_lines[list_brazil[0] + 1]]

        list_valo = [i for i,val in enumerate(all_lines) if 'Valores' in val]
        
        j = 0
        
        for i in range(4):
            name_list = ['Delivery tax', 'Shipping Currency', 'Insurance', 'Insurance Money type', 
                         'VMLE', 'Money type VMLE', 'VMLD', 'Money type VMLD']
            
            val = re.sub(r"\s+"," ",all_lines[list_valo[0] + 2 + i]).strip().split(" ")[-1]
            
            maketrans = val.maketrans
            final_val = val.translate(maketrans(',.', '.,'))
            dict1[name_list[j]] = [final_val]
            
            curr = " ".join(re.sub(r"\s+"," ",all_lines[list_valo[0] + 2 + i]).strip().split(" ")[-4:-1])
            f = open('money_type.json')
            data = json.load(f)
            
            if curr.lower() in data:
                dict1[name_list[j + 1]] = [data[curr.lower()]]
                
            j+=2
        
        """
        This is for Adica0-Check
        """
        
        dict2 = {}
        check = []
        
        list_adic = [i for i,val in enumerate(all_lines) if 'Quantidade de Adições:' in val]

        dict2['DI'] = [int(di)]
        dict2['Total additions/DI'] = [int(all_lines[list_adic[0]].split(":")[1].strip())]
        dict2['Total additions/Table'] = [int(all_lines[list_adic[0]].split(":")[1].strip())]
        
        for i in range(len(dict2['Total additions/DI'])):
            if dict2['Total additions/DI'][i] == dict2['Total additions/Table'][i]:
                check.append(int('1'))
        
        dict2['check'] = check
        
        
        """
        This is for Raw_Data
        """
        
        dict4 = {}
        que = []
        qt_vucv = []
        qt_iq = []
        qt_mt = []
        pn = []
        lis_adi = []
        lis_adi1 = []
        total_vucv = []
        round_total_vucv = []
        final_total_vucv = [] 
        
        new_list = [i for i,val in enumerate(all_lines) if val=='Descrição Detalhada da Mercadoria' or 
                    val == 'Imposto de Importação']

        for i in range(1, len(new_list), 2):
            a = all_lines[new_list[i-1] + 1:new_list[i]]
            
            for i in range(len(a)):
                answer = {}
                if 'Qtde:' in a[i]:
                    k = re.sub(r"\s+"," ",a[i]).split(" ")
                    qt_iq.append((int(k[1].replace(",",""))/100000))
                    qt_vucv.append(float(k[4].replace(",",".")))
                    
                    money_type = " ".join(k[-3:])
                    f = open('money_type.json')
                    data = json.load(f)
                        
                    if money_type.lower() in data:
                        qt_mt.append(data[money_type.lower()])
        
        for i in range(1, len(new_list), 2):
            a = all_lines[new_list[i-1] + 1:new_list[i]]
            
            for i in range(len(a)):
                if 'P/N:' in a[i]:
                    k = re.sub(r"\s+"," ",a[i]).split(":")[-1].strip().replace(")","")
                    pn.append(k)
        
        new_list_adi = [i for i,val in enumerate(all_lines) if 'Declaração:' in val]
        
        for new in new_list_adi:
            if 'Adição:' in (all_lines[new+1]):
                adicao = all_lines[new+1].split(" ")[-2].lstrip('0')
                lis_adi1.append(adicao)
                for i in range(int(Decimal(len(pn)/2).quantize(0, ROUND_HALF_UP))):
                    lis_adi.append(adicao)

        for i in range(len(qt_vucv)):
            total_vucv.append(round((qt_vucv[i] * qt_iq[i]), 2))

        for i in range(len(total_vucv)):
            round_total_vucv.append(int(Decimal(total_vucv[i]).quantize(0, ROUND_HALF_UP)))
   
        for i in range(len(round_total_vucv)):
            final_total_vucv.append(qt_mt[i] + " " +str(round_total_vucv[i]))
        
        dict4['DI'] = [int(di)] * len(pn)
        dict4['Adicao'] = lis_adi
        dict4['PN'] = pn
        dict4['VUCV'] = qt_vucv
        dict4['Money tipe'] = qt_mt
        dict4['Item Quantity'] = qt_iq
        dict4['Total VUCV'] = total_vucv
        dict4['VMLD (split)'] = final_total_vucv
        
        """
        This is for VCMV-Check
        """
         
        dict3 = {}
        answer = []
        money_val = []
        
        new_list_ven = [i for i,val in enumerate(all_lines) if val=='Condição de Venda']
        
        for i in range(len(new_list_ven)):
            vcmv = re.sub(r"\s+"," ",all_lines[new_list_ven[i] + 2]).split(" ")[1]
            maketrans = vcmv.maketrans
            final = vcmv.translate(maketrans(',.', '.,', ' '))
            money_val.append(final)

        for i in range(len(new_list_ven)):
            money = " ".join(re.sub(r"\s+"," ",all_lines[new_list_ven[i] + 2]).split(" ")[-4:-1])
            
            if money.lower() in data:
                answer.append(data[money.lower()])
        
        dict3['DI'] = [int(di)] * len(lis_adi1)
        dict3['Addition'] = lis_adi1
        dict3['VCMV'] = money_val
        dict3['Money type'] = answer
        
        """
        This is for Summary
        """
        
        dict5 = {}
        dict_vmld = {}
        temp_list_pn = []
        temp_dict_amount = {}
        temp_list_vmld = []
        sum_value_vmld = []
        temp_list_money = []
        temp_list_amount = []
        sum_value_amount = []
        
        for ele in dict4['PN']:
            if ele in temp_list_pn:
                continue
            else:
                temp_list_pn.append(ele)

        num_money_type = len(temp_list_pn)
        temp_list_money = ['USD'] * num_money_type

        for x, y in zip(dict4['PN'], dict4['Item Quantity']):
            temp_list_amount.append((x, y), )

        for line in temp_list_amount:
            temp_dict_amount.setdefault(line[0], []).append(line[1])

        for key, value in temp_dict_amount.items():
            sum_value_amount.append(sum(temp_dict_amount[key]))

        for x, y in zip(dict4['PN'], round_total_vucv):
            temp_list_vmld.append((x, y), )

        for line in temp_list_vmld:
            dict_vmld.setdefault(line[0], []).append(line[1])

        for key, value in dict_vmld.items():
            sum_value_vmld.append('USD' + " " + str(sum(dict_vmld[key])))
            
        dict5['PN'] = temp_list_pn
        dict5['Money type'] = temp_list_money
        dict5['Total amount'] = sum_value_amount
        dict5['VMLD(split)'] = sum_value_vmld


def check_workbook(wb_path):
    if os.path.exists(wb_path):
        return openpyxl.load_workbook(wb_path)
    return openpyxl.Workbook()

check_workbook('New_DI.xlsx')

excelpath = 'New_DI.xlsx'

writer = pd.ExcelWriter(excelpath, engine='xlsxwriter')

df1 = pd.DataFrame(dict1)
df2 = pd.DataFrame(dict2)
df3 = pd.DataFrame(dict3)
df4 = pd.DataFrame(dict4)
df5 = pd.DataFrame(dict5)

df4.to_excel(writer, sheet_name = 'Raw_Data')
df1.to_excel(writer, sheet_name='VMLD')
df5.to_excel(writer, sheet_name = 'Summary')
df3.to_excel(writer, sheet_name='VCMV-Check')
df2.to_excel(writer, sheet_name = 'Adica0-Check')

writer.save()