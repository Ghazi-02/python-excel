from openpyxl import Workbook,load_workbook


wb=load_workbook("tyres2.xlsx")

def minPriceSorter(wb):
    """Takes in a workbook of tyres and returns a dictionary sorting
     through each tyre to find the cheapest ones.
     Can be made more generic"""
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
    print("DEBUG:",f"{dict}\n")
    print("DEBUG:",len(dict))
    return dict

def cheapestTyreTable(dictionary,file):
    """Turns a dictionary into an excel table. File should be the output destination of the table.
    Currenty used in conjunction with minPriceSorter().
    Can be made more generic"""
    y = 1
    while y < len(dictionary):
        wb = load_workbook(file)
        ws = wb.active
        for key in dictionary:
            ws[f"A{y}"].value = key
            ws[f"B{y}"].value = dictionary[key][0]
            ws[f"C{y}"].value = dictionary[key][1]
            wb.save(file)
            y+=1
