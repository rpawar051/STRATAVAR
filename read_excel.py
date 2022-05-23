# Program to extract a particular row value
import xlrd, datetime
import json

loc = ("ToParse_Python .xlsx")
 
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)   # assign sheet index position
 


#print(sheet.nrows) # Extracting number of rows
new_list = list()
def top_column(start, end, columns=[1,2]):
    d = dict()  # create a empty dict
    # rows caan be access using these for loop
    for i in range(start, end):
        # check the value is empty return then skip these rows
        if sheet.cell_value(i, columns[0])=='' and sheet.cell_value(i, columns[1])=='':
            continue
        # check the value is return in date formate then convert in perfect date formate
        elif sheet.cell_value(i, columns[0])=='Date':
            date = datetime.datetime(*xlrd.xldate_as_tuple(sheet.cell_value(rowx=1, colx=5), wb.datemode))
            d[sheet.cell_value(i, columns[0])] = str(date)
        else:
            # add element in dict with passing key and value pairs
            d[sheet.cell_value(i, columns[0])] = sheet.cell_value(i, columns[1])
        
        return d  # return dict

# pass row number 1 to 6 data convert in dictinary format
for i in range(1, 6, 2):
    # call create_dict() function with pass rows value return dict append in list
    new_list.append(top_column(i, i+1,[1,2]))

# call create_dict() function with pass rows value return dict append in list          
new_list.append(top_column(1, 2,[4,5]))   # date value added


def create_dict(start, end):
    new_dict = dict()  # create empty dict
    # columns value get these for loop
    for i in range(sheet.ncols):   
        # rows value get these for loop 
        for j in range(start,end):
            # check the value is empty return then skip these rows
            if sheet.cell_value(8, i)=='' and sheet.cell_value(j, i)=='':
                continue
            else:
                # add element in dict with passing key and value pairs
                new_dict[sheet.cell_value(8, i)] = sheet.cell_value(j, i)
    # return created dictinary
    return new_dict      

# pass row number 9 to 11 data convert in dictinary format
for i in range(9, 11):  
    # call create_dict() function with pass rows value return dict append in list
    new_list.append(create_dict(i, i+1)) 



json_object = json.dumps(new_list)  #  convert python dict to json objects
print(json_object)