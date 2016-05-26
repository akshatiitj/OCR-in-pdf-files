# -*- coding: utf-8 -*-
"""
Created on Sun May 15 11:30:51 2016

@author: akshatshah
"""
import list2excel

def find_index_2(data, ll):
    while True:
        if data[ll] == 'B':
            if data[ll+1:ll+8] == 'alance*':
                return ll+8
        ll+=1

def convert_one_row(substr, ii):
    date = substr[ii   : ii+8].strip()
    des  = substr[ii+11: ii+53].strip()
    debt = substr[ii+80: ii+99].strip()
    ref  = substr[ii+53: ii+72].strip()
    cred = substr[ii+99:ii+118].strip()
    blnc = substr[ii+118:ii+135].strip()
    ii += 135
    if substr[ii] == ' ':
        #In case of 'blnc'               STATEMENT SUM
        if substr[ii:ii+28] == '               STATEMENT SUM':
            return [date, des, ref, debt, cred, blnc, -1]
        ref_pt = ii
        while True:
            if substr[ii] == '/' and substr[ii+3] == '/' and substr[ii+6:ii+9] == '   ':
                break
            #In case of 'des_2'               STATEMENT SUM            
            if substr[ii:ii+28] == '               STATEMENT SUM':
                des = des + ' ' + substr[ref_pt+11:ii]
                return [date, des, ref, debt, cred, blnc, -1]
            if substr[ii:ii+9] == 'Page No.:':
                des = des + ' ' + substr[ref_pt+11:ii]
                return [date, des, ref, debt, cred, blnc, ii]
            ii+=1
        des = des + ' ' + substr[ref_pt+11:ii-2]
        return [date, des, ref, debt, cred, blnc, ii-2]
    return [date, des, ref, debt, cred, blnc, ii]

def format_row(item, row_no):
    date = '20' + item[0][6:8] + '-' + item[0][3:5] + '-'+ item[0][0:2] + ' 00:00:00'
    #excel_format[mm][1] =   // Modify item[1] (date) and 6throw it in.
    des = item[1]
    ref = item[2]
    if item[3] == '':
        judg = 'C'
    else:
        judg = 'D'
    debt = ''
    for a in item[3]:
        if a != ',':
            debt += a
    cred = ''
    for a in item[4]:
        if a != ',':
            cred += a
    blnc = ''
    for a in item[5]:
        if a != ',':
            blnc += a
    name = ''
    if item[1][0:4] == 'NEFT':
        count = 0        
        for a in item[1]:
            if a == '-':
                count += 1
            if count == 2 and a != '-':
                name = name + a
    return [row_no, date, des, ref, judg, debt, cred, blnc, name]

def main():
    with open("output.txt") as f:
        filedata=f.readlines()
    
    data = filedata[0]
    
    ii = 0
    while True:
        if data[ii] == 'R':
            if data[ii+1:ii+7] == 'EGULAR':
                ii = ii+7
                break
        ii+=1
    
    excel_arr = []
    excel_arr.append(convert_one_row(data, ii))
    kk = 0
    while True:
        excel_arr.append(convert_one_row(data, excel_arr[kk][6]))
        if data[excel_arr[kk+1][6]] == 'P':
            excel_arr[kk+1][6] = find_index_2(data, excel_arr[kk+1][6])
        if excel_arr[kk+1][6] == -1:
            break
        kk+=1
        
    excel_format = []
    excel_format.append([0, 'Date', 'Description', 'Reference', 'D/C', 'Debit', 'Credit', 'Balance', 'Name'])
    
    count = 0
    for item in excel_arr:
        count += 1
        excel_format.append(format_row(item, count))
        
    list2excel.py2excel(excel_format)
    
main()