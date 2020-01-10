import pdfplumber, os, re, xlwt, xlrd


def Convert(string):
    li = list(string.split(" "))
    return li

def right(s, amount):
    return s[-amount:]

#create list of files to parse
pdfFiles = []
for filename in os.listdir(path = 'C:\\Users\\jdeasis002\\Desktop\\Python\\Bank Stmnts'):
    if filename.endswith('.pdf'):
        pdfFiles.append(filename)

prev_bal = {}
dep = {}
check = {}
end_bal = {}
#wb = xlwt.workbook()
#ws = wb.add_sheet('New Sheet')

for files in pdfFiles:
    #dict = {}
    date = files.replace('.pdf','')
    pdf_doc = pdfplumber.open('C:\\Users\\jdeasis002\\Desktop\\TS Projects\\Project Future\\FOS Bolt-on\\{}'.format(files))
    for page in pdf_doc.pages:
        if page.page_number == 1:
            pagetext = Convert(page.extract_text())

            DepIndex = pagetext.index("Deposits/Credits") + 1
            CheckIndex = pagetext.index("Checks/Debits") + 1

            right_list = []
            for i in pagetext:
                if i == '78633\nACCOUNT':
                    i = "78633"
                right_list.append(str(right(i,7)))
            end_bal_index = right_list.index('ACCOUNT')


            prev_bal_index = pagetext.index("Number") -1

            #print(right_list)
            prev_bal[date] = pagetext[prev_bal_index]
            dep[date] = pagetext[DepIndex]
            check[date] = pagetext[CheckIndex]
            end_bal[date] = pagetext[end_bal_index]




print(prev_bal)
print(dep)
print(check)
print(end_bal)

workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('test sheet')

row = 0
col = 0

for key, value in prev_bal.items():
    col += 1
    worksheet.write(row, col, key)
    worksheet.write(row +1, col, value)

col=0
for value in dep.values():
    col+=1
    worksheet.write(row + 2, col, value)
col=0
for value in check.values():
    col+=1
    worksheet.write(row + 3, col, value)
col=0
for value in end_bal.values():
    col+=1
    worksheet.write(row + 4, col, value)

workbook.save('cashproof.xls')
