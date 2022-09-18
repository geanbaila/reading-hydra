import xlsxwriter
import operator
from os import walk
import os

class Index:
    row_magnitude = { '5':'', '10':'', '15':'', '20':'', '25':'', '30':'', '35':'', '40':'', '45':'', '50':'', '55':'' }
    column_base = 17
    def getNextLetter(self, value):
        response = '0'
        for key, magnitude in self.row_magnitude.items():
            if magnitude == value:
                response = key
            
        if (response == '0'):
            available_magnitude = sorted(self.row_magnitude.items(), key=operator.itemgetter(1))[0]
            if (available_magnitude[1] == ''):
                response = available_magnitude[0]
                self.row_magnitude.update({response: value})
        return int(response)
    
    def writeHeaders(self, sheet):
        for key, value in self.row_magnitude.items():
            sheet.write(0, self.column_base + int(key), value)
        
    def run(self):
        f = []
        input_path = input('enter your path directory :\n')
        valid_extension = '.gse'

        dir = 0
        for (root, dirs, files) in os.walk(input_path):
            if os.path.isdir(root):
                if files:
                    dir+=1
                    myxlsx = xlsxwriter.Workbook('file'+str(dir)+'.xlsx')
                    sheet = myxlsx.add_worksheet()
                    row = 1
                    for file in files:
                        filename, filename_ext = os.path.splitext(file)
                        if filename_ext.lower() == valid_extension:
                            blocked_a = False
                            blocked_b = False
                            blocked_c = False
                            with open(os.path.join(root, file)) as xfile:
                                espacios = 0
                                linea = 0
                                print(30*"=", file, 30*"=")
                                for line in xfile:
                                    informacion = line.rstrip()
                                    if informacion == '' and espacios == 0:
                                        espacios+= 1
                                        linea = 1
                                        blocked_a = True
                                        continue
                                    elif informacion =='' and espacios == 1:
                                        espacios+=1
                                        linea = 1
                                        blocked_a = False
                                        blocked_b = True
                                        continue
                                    elif informacion =='' and espacios == 2:
                                        espacios+=1
                                        linea = 1
                                        blocked_a = False
                                        blocked_b = False
                                        blocked_c = True
                                        continue
                                    else:
                                        linea+=1
                                        if blocked_a and linea == 3:
                                            col = 0
                                            row_line = line.split()
                                            for l in row_line:
                                                sheet.write(row, col, l)
                                                col+=1
                                            row+=1
                                        elif blocked_b and linea >= 3:
                                            row_line = line.split()                                            
                                            new_col = self.getNextLetter(row_line[0]) 
                                            if (new_col != 0):
                                                sheet.write(row, self.column_base + new_col, row_line[1])
                                                new_col+=1
                                                sheet.write(row, self.column_base + new_col, row_line[2])
                                                new_col+=1
                                                sheet.write(row, self.column_base + new_col, row_line[3])
                                                new_col+=1
                                                sheet.write(row, self.column_base + new_col, row_line[4])
                                                new_col+=1
                                                sheet.write(row, self.column_base + new_col, row_line[5])
                                                new_col+=1
                                        else:
                                            continue
                    self.writeHeaders(sheet)
                    myxlsx.close()
        
index = Index()
index.run()