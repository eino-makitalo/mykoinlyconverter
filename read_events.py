#--*-- coding: utf-8 --*--

import csv
import datetime
import pytz
import shutil
import os
from openpyxl import Workbook,load_workbook
from openpyxl.utils import get_column_letter

hmap={}

WORKDIR=r"c:\temp"
from settings import WORKDIR  # settings.py is inl

FIFO_TEMPLATE="verohallinto_-fifo-laskuri_versio-1.1.xlsm" 

CURRENCY_COL_IN_TEMPLATE=8   # column H  after text "Laskelmassa käytetty virtuaalivaluutta"
CURRNECY_ROW_IN_TEMPLATE=9   # row where currency is saved (just information)



KOINLYTRANSACTIONFILES_TO_READ = ("Koinly_2021.csv","Koinly_2022.csv")


def ExcelForCurrency(currency):
    return "%s_FIFO.xlsm" % (currency,)

hki = pytz.timezone("Europe/Helsinki")

BOOK_CURRENCY="EUR"  # bookkeeping currency

F_DATE="Date"
F_TYPE="Type"
F_LABEL="Label"
F_SWALLET="Sending Wallet"
F_SAMOUNT="Sent Amount" # in crypto
F_SCURRENCY="Sent Currency"
F_SCOST="Sent Cost Basis" # in bookkeeping currency
F_RWALLET="Receiving Wallet"
F_RAMOUNT="Received Amount" # in crypto
F_RCURRENCY="Received Currency"
F_RCOST="Received Cost Basis" # in bookkeeping currency
F_FEE="Fee Amount"
F_FCURRENCY="Fee Currency"
F_GAIN="Gain (%s)" % (BOOK_CURRENCY,)
F_NETVAL="Net Value (%s)" % (BOOK_CURRENCY,)
F_FEEVAK="Fee Value (%s)" % (BOOK_CURRENCY,)
F_TXSRC="TxSrc"
F_TXDEST="TxDest"
F_TXHASH="TxHash"
F_DESC="Description"

# Date,Type,Label,Sending Wallet,Sent Amount,
# Sent Currency,Sent Cost Basis,Receiving Wallet,Received Amount,Received Currency,
# Received Cost Basis,Fee Amount,Fee Currency,
# Gain (EUR),Net Value (EUR),Fee Value (EUR),TxSrc,TxDest,TxHash,Description

TIMEZONE="Europe/Helsinki"  # just used to separate months for bookkeeping

FIFO_ABS=os.path.join(WORKDIR,FIFO_TEMPLATE)
if not os.path.exists(FIFO_ABS):
    raise AssertionError("""We cant file file %s  
    in directory %s
    download it from\n https://www.vero.fi/tietoa-verohallinnosta/yhteystiedot-ja-asiointi/verohallinnon_laskuri/fifo-laskuri/"""%(FIFO_TEMPLATE,WORKDIR))

import os

TYPES=set()
Currencies=set()
Wallets=set()

first_csv_file=True
for file1 in KOINLYTRANSACTIONFILES_TO_READ:
    absfile1=os.path.join(WORKDIR,file1)
    skip_until_header=True
    with open(absfile1,"r") as csvfile:
        for row in csv.reader(csvfile,dialect='excel'):        
            if skip_until_header:
                if len(row)==20 and row[0]=='Date':
                    skip_until_header=False
                    # we read this from first csv only - should be same mapping
                    ind=0
                    for cell in row:
                        if first_csv_file:
                            hmap[cell]=ind
                        else:
                            if (ind != hmap[cell]):
                                raise AssertionError("Transaction files from Koinly has different format - can't continue")
                        ind=ind+1                    
            else:
                utctime=pytz.utc.localize(datetime.datetime.strptime(row[0],'%Y-%m-%d %H:%M:%S UTC'))
                #print(utctime)
                x1=hki.normalize(utctime)
                print(utctime, x1)
                TYPES.add(row[hmap[F_TYPE]])
                if row[hmap[F_SCURRENCY]]:
                    Currencies.add(row[hmap[F_SCURRENCY]])
                if row[hmap[F_RCURRENCY]]:
                    Currencies.add(row[hmap[F_RCURRENCY]])
                if row[hmap[F_RWALLET]]:
                    Wallets.add(row[hmap[F_RWALLET]])
                if row[hmap[F_SWALLET]]:
                    Wallets.add(row[hmap[F_SWALLET]])
                    
# We will create one Tax calculation Excel for every wallet
# Please, if you have same currency USDC in multiple blockchain you should consider handling them separately ?

EXCELMAP={}

import tempfile

def saveexcel(wb,absfile):
    dir1,file1=os.path.split(absfile)
    fn1,ext1 = os.path.splitext(file1)    
    fdtemp,absfiletemp=tempfile.mkstemp(suffix=ext1,prefix=fn1)
    os.close(fdtemp)
    wb.save(absfiletemp)
    shutil.copy2(absfiletemp,absfile)
    os.remove(absfiletemp)
    

for curr in Currencies:
    if curr!="AVAX":
        continue
        
    absfile2=os.path.join(WORKDIR,ExcelForCurrency(curr))
    EXCELMAP[curr]=absfile2
    if os.path.exists(absfile2):
        print(f"We have already excel for {curr}")
    else:
        print(f"We create new exec for {curr} from template")
        shutil.copy2(os.path.join(WORKDIR,FIFO_TEMPLATE), absfile2)
    wb=load_workbook(absfile2,keep_vba=True)
    first_sheet=wb.worksheets[0]
    curr2=first_sheet.cell(CURRNECY_ROW_IN_TEMPLATE,CURRENCY_COL_IN_TEMPLATE).value
    if (curr!=curr2):
        assert(curr2==None)
        print(f"We set excel right currency {curr}")
        first_sheet.cell(CURRNECY_ROW_IN_TEMPLATE,CURRENCY_COL_IN_TEMPLATE).value=curr
        saveexcel(wb, absfile2)
    print(f"Currency {curr} is set in excel workbook")
        
        

print(TYPES)
print(Currencies)
print(Wallets)