from attr import resolve_types
import pdfquery
import calendar
from pikepdf import Pdf
import os
import sys


pagesInfo = []
uniqueClients = []

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'





#pdf
def splitDocument():
    print(bcolors.OKGREEN + 'Great' + bcolors.ENDC)
    print('')
    print('We will now split and rename the document...')
    folder = documentName.replace('.pdf','/')
    if not os.path.isdir(folder):
        os.mkdir(folder)
    fullFile = Pdf.open(documentName)
    for client in uniqueClients:
        file = Pdf.new()
        for p in range(int(client['start_page'])-1,int(client['end_page'])):
            file.pages.append(fullFile.pages[p])
        filename = '%s, %s, %s %s, %s.pdf' % (client['client'], client['currency'], client['month'], client['year'], client['type'])
        print("saving "+ bcolors.UNDERLINE + filename + bcolors.ENDC)  
        file.save(folder + filename)
    print(bcolors.WARNING + '[!]'+bcolors.ENDC+' files are saved to the ' + bcolors.BOLD + folder + bcolors.ENDC + ' folder')      
    print('')  
    print(bcolors.OKGREEN + 'This has been an EXCEEDINGLY successful day' + bcolors.ENDC)
    print('Good bye')   

def pagesplitTest():
    fullFile = Pdf.open(documentName)
    file = Pdf.new()
    file.pages.append(fullFile.pages[1])
    file.save('test.pdf')

def getItemOnPage(pageSelector, itemSelector):
    return pdf.extract([
        ('with_parent',pageSelector,),
        ('value', itemSelector) # only matches elements on page 1
    ])['value']

def locationMethod(pageNo):
    pdf.load(pageNo)
    #pageSelector = 'LTPage[page_index="' + str(pageNo) + '"] '
    to_label = pdf.pq('LTTextLineHorizontal:contains("TO"):not(:contains("TOTAL"))')
    date_label = pdf.pq('LTTextLineHorizontal:contains("DATE ")')
    #date_label = pdf.pq('LTTextLineHorizontal:overlaps_bbox("%s, %s, %s, %s")' % (425, 625, 510, 635))
    if len(date_label ) > 1:
        date_label = date_label
    date_text = date_label.text()
    (monthInt, year) = getMonthAndYear(date_text)
    #is_statement = pdf.pq('LTTextBoxHorizontal:contains("Statement")') != ''
    is_statement = pdf.pq('LTTextBoxHorizontal:contains("INVOICE"):nth-child(1)') == []
    client_name = getBelowText(to_label)
    currency = 'CAD'
    usFindIndex = client_name.find('US -')
    if (usFindIndex > -1):
        client_name = client_name[5:]
        currency = 'USD'
    
    return {
        'start_page': pageNo+1,
        'end_page' : pageNo+1,
        'client': client_name.replace('.',''),
        'month' : calendar.month_name[monthInt],
        'year': year,
        'type': 'statement' if is_statement else 'invoice',
        'currency': currency
    }

def getBelowText(label):
    left_corner = float(label.attr('x0'))
    bottom_corner = float(label.attr('y0'))
    return pdf.pq('LTTextLineHorizontal:overlaps_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner-12, left_corner+100, bottom_corner-2)).text()

def getMonthAndYear(dateStr):
    monthInt = 1
    year = ''

    if len(dateStr) > 14 and dateStr[0:4] == "DATE":
        monthInt = int(dateStr[6:7])
        year = dateStr[11:15]

    return (monthInt, year)

def getPageCount():
    value = pdf.doc.catalog['Pages'].resolve()['Count']  
    return value

def loopPages():
    pageCount = getPageCount()
    lastData = locationMethod(0)
    for i in (range(pageCount-1)):
        page = i+1
        pageData = locationMethod(page)
        if lastData['client'] == pageData['client']:
            # Another page for current customer
            lastData['end_page'] = page+1         
        else:
            # New customer, add previous to list
            uniqueClients.append(lastData)
            lastData = pageData
        if page == pageCount-1:
                uniqueClients.append(lastData)
        
    
    print("Pages: %s    | Clients: %s" % (pageCount, len(uniqueClients))) 
    

def die():
    print(bcolors.WARNING + '                            . ')
    print('                           -|- ')
    print('                            | ')
    print(bcolors.ENDC + '                        .-\'~~~`-. ')
    print('                      .\'         `. ')
    print('                      | ' + bcolors.OKCYAN + ' R  I  P ' + bcolors.ENDC + ' |   to Matts code')
    print('                      |           | ')
    print('                      |           | ')
    print(bcolors.OKGREEN + '                     \\' + bcolors.ENDC + '|           |' + bcolors.OKGREEN + '// ')
    print('^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^' + bcolors.ENDC)
    exit()

def startupHeader():
    print('Startup sequence...')
    print(bcolors.ENDC + ' _____   ____  ____   ____________  ' + bcolors.OKCYAN + '         _   ___   _ ')
    print(bcolors.ENDC + '|  __ \ / __ \|  _ \ / __ \__   __| ' + bcolors.OKCYAN + '   /\   | \ | | \ | |   /\     ')
    print(bcolors.ENDC + '| |__) | |  | | |_) | |  | | | |    ' + bcolors.OKCYAN + '  /  \  |  \| |  \| |  /  \    ')
    print(bcolors.ENDC + '|  _  /| |  | |  _ <| |  | | | |    ' + bcolors.OKCYAN + ' / /\ \ | . ` | . ` | / /\ \   '+ bcolors.ENDC +'v1.0.0'+bcolors.OKGREEN )
    print(bcolors.ENDC + '| | \ \| |__| | |_) | |__| | | |    ' + bcolors.OKCYAN + '/ ____ \| |\  | |\  |/ ____ \  ')
    print(bcolors.ENDC + '|_|  \_\\_____/|____/ \____/  |_|   ' + bcolors.OKCYAN + '/_/    \_\_| \_|_| \_/_/    \_\ ' + bcolors.HEADER)
    print('~ The Magical Document Reading, Grouping and Renaming system ~'+bcolors.ENDC)
    print('')

    
    

def printClientDetails():
    print('')
    print("CLIENT NAME                       | PAGES   | CUR | DATE          | TYPE" )
    print("----------------------------------|---------|-----|---------------|---------" )
    for c in uniqueClients:
        pagesStr = "%s-%s" % (c['start_page'],c['end_page'] )
        if c['start_page'] == c['end_page']:
            pagesStr = "%s" % (c['start_page'])
        print(" %s   | %s | %s | %s %s | %s" % (c['client'].ljust(30,' '),pagesStr.center(7,' '),c['currency'],c['year'],c['month'].ljust(8,' '),c['type']))
    print("----------------------------------|---------|-----|---------------|---------" )
    print("")
    key = ''
    while key != 'Y' and key != 'n':   
        key = input("Does this look good to you? [Y/n] ")
    if key == 'n':
        print(bcolors.FAIL + 'Oh no. I guess I\'ll die then.' + bcolors.ENDC)
        die()
    if key == 'Y':
        splitDocument()

# START

startupHeader()

# Check arguments
if len(sys.argv) < 2:
        noFile = True
        while noFile:
            documentName = input('Please enter the file name [q=quit]: ')
            if documentName == 'q':
                print(bcolors.FAIL + 'Bye bye have a beautiful time'+ bcolors.ENDC)
                die()
            if os.path.isfile(documentName):
                noFile = False
            else:
                print("The file '"+documentName+"' doesn't exist.")
else: 
    documentName = sys.argv[1] #.replace('.\\','')
    if not os.path.isfile(documentName):
        exit("The file '"+documentName+"' doesn't exist.")

print('Reading document: ' + bcolors.UNDERLINE + documentName + bcolors.ENDC)

#read the PDF
pdf = pdfquery.PDFQuery(documentName)

# Load the PDF for reading
pdf.load()

# Read in all data
loopPages()

# Print details and finish up
printClientDetails()