from openpyxl import Workbook,load_workbook


#create a workbook object
#wb = Workbook()

#load existing speadsheet

wb=load_workbook('tyres.xlsx')

#create an active  worksheet
ws = wb.active

"""
dic = {}
x = 1
while x < 6:
    size = ws[f"A{x}"].value
    price = ws[f"B{x}"].value
    dic.update({size :f"{price}"})
    x += 1
print(dic)
#Returns dictionary dic in form {'Tyres': 'Price', '135 40 14': '15', '155 55 16': '20', '195 55 16': '25', '185 35 18': '30'}
"""
def price_changer(file,qty,coloumn):    #arg file and coloumn must be strings 
    """ Changes value of each cell in {coloumn} by {qty} in {file} (file AKA workbook).
    Parameters file and coloumn must be strings and qty must be an integer
    
    """ 
    " Can be made more generic"                                    
    coloumn_ = ws[coloumn]
    print(coloumn)
    for cell in coloumn_:
        if type(cell.value) == str:
            cell.value = cell.value
        else:
            cell.value = cell.value + qty 
    wb.save(file)


def brand_changer(file,brand,coloumn):  #args must be string
    """ Changes  the brand of tyres to {brand} in {coloumn} in {file}. arguments must be strings"""
    "Can be made more generic"                         
    coloumn_ =ws[coloumn]
    for cell in coloumn_:
        if cell.value == "Brand":
            cell.value
        else:
            cell.value = f"{brand}"
    wb.save(file)

def sheet_creator(workbook,name): 
    """ Creates an empty sheet , arguments must be strings"""
    wb = load_workbook(workbook)
    wb.create_sheet(name)
    wb.save(workbook)

print(wb.sheetnames)
