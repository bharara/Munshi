from openpyxl import load_workbook
wb = load_workbook('m2.xlsx')

# Gets the Client List

client = wb.get_sheet_by_name('Client')

for n in range(2,88):
    client['A'+str(n)] = client['A'+str(n)].value.title()
    try:
        client['A'+str(n)] = client['A'+str(n)].value.replace('/','-')
    except:
        pass

wb.save('m2.xlsx')
