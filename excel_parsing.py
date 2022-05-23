import xlrd, datetime, json
class Excel_parsing():
    def __init__(self, file_name):
        self.loc = (file_name)
        self.wb = xlrd.open_workbook(self.loc)  # read excel file
        self.sheet = self.wb.sheet_by_index(0)
        self.new_list = list()
    
    def top_column(self, start, end, columns=[1,2]):
        self.d = dict()  # create a empty dict
        # rows caan be access using these for loop
        for i in range(start, end):
            # check the value is empty return then skip these rows
            if self.sheet.cell_value(i, columns[0])=='' and self.sheet.cell_value(i, columns[1])=='':
                continue
            # check the value is return in date formate then convert in perfect date formate
            elif self.sheet.cell_value(i, columns[0])=='Date':
                self.date = datetime.datetime(*xlrd.xldate_as_tuple(self.sheet.cell_value(rowx=1, colx=5), self.wb.datemode))
                self.d[self.sheet.cell_value(i, columns[0])] = str(self.date)
            else:
                # add element in dict with passing key and value pairs
                self.d[self.sheet.cell_value(i, columns[0])] = self.sheet.cell_value(i, columns[1])
            
            return self.d  # return dict
    
    def create_dict(self, start, end):
        self.new_dict = dict()  # create empty dict
        # columns value get these for loop
        for i in range(self.sheet.ncols):   
            # rows value get these for loop 
            for j in range(start,end):
                # check the value is empty return then skip these rows
                if self.sheet.cell_value(8, i)=='' and self.sheet.cell_value(j, i)=='':
                    continue
                else:
                    # add element in dict with passing key and value pairs
                    self.new_dict[self.sheet.cell_value(8, i)] = self.sheet.cell_value(j, i)
        # return created dictinary
        return self.new_dict      

    def create_list(self):
        # pass row number 1 to 6 data convert in dictinary format
        for i in range(1, 6, 2):
            # call create_dict() function with pass rows value return dict append in list
            self.new_list.append(self.top_column(i, i+1,[1,2]))

        # call create_dict() function with pass rows value return dict append in list          
        self.new_list.append(self.top_column(1, 2,[4,5]))   # date value added

        # pass row number 9 to 11 data convert in dictinary format
        for i in range(9, 11):  
            # call create_dict() function with pass rows value return dict append in list
            self.new_list.append(self.create_dict(i, i+1)) 

    def convert_list_to_json(self):
        return json.dumps(self.new_list) #  convert python list to json objects

ep = Excel_parsing("ToParse_Python .xlsx")  # create Excel_parsing class objects
ep.create_list()  # call create_list() methods
print(ep.convert_list_to_json())