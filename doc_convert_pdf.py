from win32com import client
word = client.DispatchEx('Word.Application')
import os
from optparse import OptionParser

if __name__=='__main__':
    parser=OptionParser(usage='%prog [options]')
    parser.add_option('-i','--in',dest='input',help='input file')
    parser.add_option('-o','--out',dest='output',help='output file')
    (options,args)=parser.parse_args()
    input=options.input
    output=options.output
    input = os.path.abspath(input)
    output = os.path.abspath(output)
    doc=word.Documents.Open(input)
    doc.SaveAs(output,FileFormat=17)
    doc.Close()