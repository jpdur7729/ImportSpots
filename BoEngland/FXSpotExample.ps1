# ------------------------------------------------------------------------------
#                     Author    : eFront
#                     Time-stamp: "2018-08-22 18:24:11 jpdur"
# ------------------------------------------------------------------------------


# Extract Examples from Bank of England
                             
                             
# ------------------------------------------           
# Generating a csv file with FX Spot
# ------------------------------------------          
$URL = "http://www.bankofengland.co.uk/boeapps/iadb/fromshowcolumns.asp?csv.x=yes&Datefrom=01/Feb/2006&Dateto=01/Oct/2007 &SeriesCodes=XUDLERS,XUDLUSS&CSVF=TN&UsingCodes=Y&VPD=Y&VFD=N"

$extractcmd = "wget """ + $URL + """ -o LogFile.txt -O FXSpotResult.csv"
$extractcmd
# ---------------------------------------------------------------------------------------------
# Store the command in a .bat file (Encoding ASCII guarantees that there is no odd character
# Execute the command
# Delete the intermediate file
# ---------------------------------------------------------------------------------------------
$extractcmd | Out-File -Encoding ASCII "./goextract.bat"
& "./goextract.bat"
rm goextract.bat

cat FXSpotResult.csv
