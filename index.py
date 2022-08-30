import xlsxwriter
from os import walk
import os

f = []
input_path = input('ingrese el directorio de lectura:\n')
valid_extension = '.gse'

directorio = 0
for (root, dirs, files) in os.walk(input_path):
    if os.path.isdir(root):
        if files:
            directorio+=1
            libro = xlsxwriter.Workbook('file'+str(directorio)+'.xlsx')
            hoja = libro.add_worksheet()
            row = 0
            for file in files:
                filename, filename_ext = os.path.splitext(file)
                if filename_ext.lower() == valid_extension:
                    blocked_a = False
                    blocked_b = False
                    blocked_c = False
                    with open(os.path.join(root, file)) as xfile:
                        espacios = 0
                        linea = 0
                        print("/////////// " + file)
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
                                col = 0
                                if blocked_a and linea == 3:
                                    # print('##'+line)
                                    row_line = line.split()
                                    row+=1
                                    for l in row_line:
                                        hoja.write(row, col, l)
                                        col+=1
                                elif blocked_b and linea >= 3:
                                    # print('**'+line)
                                    row_line = line.split()
                                    for l in row_line:
                                        hoja.write(row, col, l)
                                        col+=1
                                else:
                                    continue
                                    # print(informacion)
            libro.close()                

# /Users/geanbaila/Sites/reading-hydra/input/2020
# pip install XlsxWriter
