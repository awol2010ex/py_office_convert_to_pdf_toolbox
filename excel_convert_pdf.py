from win32com import client
word = client.DispatchEx('Excel.Application')
import os
from optparse import OptionParser

if __name__=='__main__':
    command = 'taskkill /F /IM et.exe'
    os.system(command)
    command = 'taskkill /F /IM excel.exe'
    os.system(command)
    
    parser=OptionParser(usage='%prog [options]')
    parser.add_option('-i','--in',dest='input',help='input file')
    parser.add_option('-o','--out',dest='output',help='output file')
    (options,args)=parser.parse_args()
    input=options.input
    output=options.output
    input = os.path.abspath(input)
    output = os.path.abspath(output)
    excel = client.DispatchEx("Excel.Application")
    excel.Visible = 0
    
    wb = excel.Workbooks.Open(Filename=input,ReadOnly=1)
    ws = excel.Worksheets[0]
    wb.ExportAsFixedFormat(0, str(output))
    
    wb.Close()
    excel.Quit()