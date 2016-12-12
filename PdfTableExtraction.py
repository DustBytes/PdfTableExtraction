import re
import sys
import subprocess
import xlsxwriter
from collections import OrderedDict


def extract(userFile, startPos, endPos):

    lstKey = []     # store first column (key)
    lstValue = []   # store second column (value)
    longDesc = []   # helps to extract multiline value

    flag = 0

    outFile = userFile.split('.')[0] + '.txt'
    proc = subprocess.Popen(['pdftotext', '-nopgbrk', '-layout', userFile, outFile]).wait()

    with open(outFile, 'r') as inFile:
        for line in inFile:
            if line.find(startPos) >= 0:
                flag = 1

            if flag == 1:
                # multiline values extraction
                if line == "\n":

                    # check if last Key has empty value, move to next page if empty
                    if not longDesc:
                        if lstKey[-1] == endPos:
                            flag = 0

                    if longDesc:
                        if lstKey[-1] == endPos:
                            flag = 0
                        lstValue.append("".join(longDesc))
                        longDesc = []
                    continue

                elif re.match(r'\s', line):
                    longDesc.append(line.strip())

                # Extract each Key, value
                elif re.match(r'\w', line):
                    new_line = line.split(':')
                    category = new_line[0]
                    catValue = new_line[1]

                    if longDesc:
                        longDesc.append(catValue.strip())

                    lstKey.append(new_line[0])
                    if not longDesc:
                        lstValue.append(new_line[1].strip())

    lstKey = list(OrderedDict.fromkeys(lstKey))     # remove duplicates

    lstValue = list(['-' if val is '' else val for val in lstValue])        # Put - to cells without value
    return lstKey, lstValue


## ------------------
## Xlxs Writing
## ------------------


def xlsWriting(lstKey, lstValue):

    workbook = xlsxwriter.Workbook(userFile.split('.')[0]+'.xlsx')
    worksheet = workbook.add_worksheet()

    format1 = workbook.add_format()
    format1.set_align('center')
    format1.set_align('vcenter')
    format1.set_bold()

    format2 = workbook.add_format()
    format2.set_align('center')
    format2.set_align('vcenter')

    row = 1
    col = 0

    for val in lstKey:
        worksheet.write(0, col, val, format1)
        col += 1

    col = 0
    for val in lstValue:
        worksheet.write(row, col,  val, format2)
        col += 1
        if col == len(lstKey):
            col = 0
            row += 1

    workbook.close()


def usage():
    print('Invalid Arguments')
    print('Usage: python3 PdfTableExtraction.py <pdf> <Starting Key of table> <End Key of Table> \n'
          'eg: python3 PdfTableExtraction.py ALL.pdf Disease\ Description Notes')


if __name__ == '__main__':
    if len(sys.argv) != 4:
        usage()
        exit(0)
    else:
        userFile = sys.argv[1]
        startPos = sys.argv[2]
        endPos = sys.argv[3]

        lstKey, lstValue = extract(userFile, startPos, endPos)
        xlsWriting(lstKey, lstValue)
        print('Finished')