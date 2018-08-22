# ------------------------------------------------------------------------------
#                     Author    : eFront
#                     Time-stamp: "2018-08-22 17:36:53 jpdur"
# ------------------------------------------------------------------------------


# Extract Examples from Bank of England


# -------------------------------
# 1 Generating a .xls spreadsheet
# -------------------------------
$URL = "http://www.bankofengland.co.uk/boeapps/iadb/fromshowcolumns.asp?excel97.x=yes&Datefrom=01/Feb/2006&Dateto=01/Oct/2007 &SeriesCodes=LPMAUZI,LPMAVAA&UsingCodes=Y&VPD=Y&VFD=N"

$extractcmd = "wget """ + $URL + """ -o LogFile.txt -O Test.xls"
$extractcmd
# ---------------------------------------------------------------------------------------------
# Store the command in a .bat file (Encoding ASCII guarantees that there is no odd character
# Execute the command
# Delete the intermediate file
# ---------------------------------------------------------------------------------------------
$extractcmd | Out-File -Encoding ASCII "./goextract.bat"
& "./goextract.bat"
rm goextract.bat

# ------------------------
# 2 Generating a xml file
# ------------------------
$URL = "http://www.bankofengland.co.uk/boeapps/iadb/fromshowcolumns.asp?CodeVer=new&xml.x=yes&Datefrom=01/Feb/2006&Dateto=01/Oct/2007 &SeriesCodes=LPMAUZI,LPMAVAA&VPD=Y&VFD=N&VUD=A&VUDdate=01/Aug/2007&Omit=-A2-B3-D"

$extractcmd = "wget """ + $URL + """ -o LogFile.txt -O Result.xml"
$extractcmd
# ---------------------------------------------------------------------------------------------
# Store the command in a .bat file (Encoding ASCII guarantees that there is no odd character
# Execute the command
# Delete the intermediate file
# ---------------------------------------------------------------------------------------------
$extractcmd | Out-File -Encoding ASCII "./goextract.bat"
& "./goextract.bat"
rm goextract.bat

cat Result.xml


# ------------------------
# 3 Generating a csv file
# ------------------------
$URL = "http://www.bankofengland.co.uk/boeapps/iadb/fromshowcolumns.asp?csv.x=yes&Datefrom=01/Feb/2006&Dateto=01/Oct/2007 &SeriesCodes=LPMAUZI,LPMAVAA&CSVF=TN&UsingCodes=Y&VPD=Y&VFD=N"

$extractcmd = "wget """ + $URL + """ -o LogFile.txt -O Result.csv"
$extractcmd
# ---------------------------------------------------------------------------------------------
# Store the command in a .bat file (Encoding ASCII guarantees that there is no odd character
# Execute the command
# Delete the intermediate file
# ---------------------------------------------------------------------------------------------
$extractcmd | Out-File -Encoding ASCII "./goextract.bat"
& "./goextract.bat"
rm goextract.bat

cat Result.csv
