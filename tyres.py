from openpyxl import Workbook,load_workbook


wb=load_workbook("tyres2.xlsx")
dict = {}
for ws in wb:
    print(ws.title)
    column_=ws['A']
    x = 1
    while x <= len(column_):
        size = ws[f"A{x}".replace(" ", "")].value
        price = ws[f"B{x}".replace(" ", "")].value
        brand = ws[f"C{x}".replace(" ", "")].value
        if size in dict:
            if dict[size][0] <= price:
                x += 1
            else:
                dict[f"{size}"]=[price,f"{brand}"]
        elif size is None or price is None:
            x += 1
        else:   
            dict.update({size :[price, f"{brand}" ]})  
            x += 1
print(f"{dict}\n")
print(len(dict))

y = 1
while y < len(dict):
    wb2 = load_workbook("sorted.xlsx")
    ws2 = wb2.active
    for key in dict:
        ws2[f"A{y}"].value = key
        ws2[f"B{y}"].value = dict[key][0]
        ws2[f"C{y}"].value = dict[key][1]
        wb2.save("sorted.xlsx")
        y+=1