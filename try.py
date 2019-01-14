import os
os.system('python basic.py')

##from openpyxl import load_workbook
##from flask import Flask, render_template, request
##import webbrowser
##from datetime import datetime
##
##wb = load_workbook(filename='Munshi.xlsx')
##data = wb.get_sheet_by_name('Data')
##
##app = Flask(__name__)
##
###User Report Shown
##@app.route('/report/user/<name>')
##def user2(name):
##
##    wb = load_workbook(filename='Munshi.xlsx', data_only=True)
##    client = wb.get_sheet_by_name('Client')
##    cid = clist.index(name) + 2
## 
##    cd1 = client['B'+str(cid)].value
##    cd2 = client['C'+str(cid)].value
##    cd3 = client['D'+str(cid)].value
##    cd4 = client['E'+str(cid)].value
##    cd5 = client['F'+str(cid)].value
##    credit = client['G'+str(cid)].value
##    debit = client['H'+str(cid)].value
##
##    name = name.replace('%20',' ')
##
##    outlist= []
##    n=1
##    for val in data['A']:
##        if val.value == name:
##        	try:
##        		hall = data['B'+str(n)].value.strftime("%d-%B-%Y")
##        	except:
##        		hall = datetime.strptime(data['B'+str(n)].value,"%d-%m-%y").strftime("%d-%B-%Y")
##        	
##        	outlist.append([hall, data['C'+str(n)].value, data['D'+str(n)].value, data['E'+str(n)].value, data['F'+str(n)].value, data['G'+str(n)].value])
##        n+=1
##
##    return render_template('user2.html', outlist=outlist, in_detail1=in_detail1, in_detail2 = in_detail2, in_detail3 = in_detail3,
##                           c_detail1 = c_detail1, c_detail2 = c_detail2, c_detail3 = c_detail3, c_detail4 = c_detail4, c_detail5 = c_detail5,
##                           cd1=cd1, cd2=cd2, cd3=cd3, cd4=cd4, cd5=cd5, credit=credit, debit=debit, name=name)
##
### Autodebug and Run the app
##if __name__ == '__main__':
##	webbrowser.open('http://127.0.0.1:5000', new=2, autoraise=True)
##	app.run(debug=False)
