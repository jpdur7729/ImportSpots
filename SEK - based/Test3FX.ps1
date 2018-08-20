# ------------------------------------------------------------------------------
#                     Author    : eFront-Mastek
#                     Time-stamp: "2018-05-09 14:52:53 jpdur"
# ------------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# The received parameter is the directory where the script is to be executed
# if nothing is received then let's use the directory where the script is
# found .... as a default
# ---------------------------------------------------------------------------
param(
	  [Parameter(Mandatory=$false)] [string] $Exec_Dir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
)

# -----------------------------------------------------------------------------------------------------
# Move to the desired Directory i.e the directory where FrontCmd.exe and config files can be found
# This is based on the parameter received
# -----------------------------------------------------------------------------------------------------
cd ($Exec_Dir)        
                      
# Configuration file to get all the corresponding parameters
. ./Configuration.ps1                                                                              
                                                                                                   
# We built the request by default for today's rate and all available symbols
# $request = "2017-12-31?access_key=eca17521f4e211d09ab357c6cd9585dc&base=EUR&symbols=USD,CAD,EUR,GBP"
$request = (Get-Date -UFormat "%Y-%m-%d") + "?access_key=eca17521f4e211d09ab357c6cd9585dc&base=EUR"

# ~~~~~~~~~~~~~~~~~~~~~~ Start Extract ~~~~~~~~~~~~~~~~~~~~~~~~~~
# Extract the FX rates from the data.fixer.io Web Site which generated json
$extractcmd = "wget ""http://data.fixer.io/api/"+ $request + """ -o LogFile.txt"

# We delete the files that we will use for the extract
rm Result.json -ErrorAction SilentlyContinue

# Debug Check extract Command
# $extractcmd

# ---------------------------------------------------------------------------------------------
# Store the command in a .bat file (Encoding ASCII guarantees that there is no odd character
# Execute the command
# Delete the intermediate file
# ---------------------------------------------------------------------------------------------
$extractcmd | Out-File -Encoding ASCII "./goextract.bat"
& "./goextract.bat"
rm goextract.bat

# We are renaming the result file with a Standard Name
mv ("./" + $request) Result.json
# ~~~~~~~~~~~~~~~~~~~~~~ Fin Extract ~~~~~~~~~~~~~~~~~~~~~~~~~~

# Read the json file as produced in the 1st part
$FxRates = Get-Content -Raw -Path ./Result.json | ConvertFrom-Json

# Display the data as found
$HashFxRates = $FxRates.rates
$BaseCCy     = $FxRates.base
$FxDate      = $FxRates.date
$FxDateasDate = [datetime]::ParseExact($FxDate,"yyyy-MM-dd", $null)

# # Debug to check date
# $FxDate
# $FxDateasDate.DateTime
# $FXDateasDate.ToString($DateFormat)

# # Way to convert/read the timestamp provided
# # https://stackoverflow.com/questions/10781697/convert-unix-time-with-powershell
# [datetime]$origin = '1970-01-01 00:00:00'
# $whatIWant = $origin.AddSeconds($FxRates.timestamp)
# $whatIWant

# -------------------------------------------------------------------------------------------------------------
# Try to process the data which is not exactly an HashTable but something close as per the article below
# https://stackoverflow.com/questions/22002748/hashtables-from-convertfrom-json-have-different-type-from-powershells-built-in-h
# The contents is adapated to the case
# -------------------------------------------------------------------------------------------------------------
# # $HashFxRates.root | select * | ft -AutoSize # Does mot work
# $HashFxRates | select * | ft -AutoSize        # Works

# SilentlyContinue does not give any error message is the file is not found...
rm  ./FXrate.csv -ErrorAction SilentlyContinue

# # The 3 elements are found in configuration
# $NewHeaderImport = 0
# $CSVSep     = ","
# $DateFormat = "dd/MM/yyyy"      #Format to be used for European Date

$OutputCSV  = 'ef$class,ef$subclass,ef$col,ef$col,ef$col,ef$col'

# The type of header to be used depends on the system configuration
if ($NewHeaderImport -eq 0) {
    $OutputCSV += "`n"+ "CurrencyRates,Standard,Rates.Reference date,Rates.Destination Currency,Rates.Source Currency,Rates.Rate"
}
else {
    $OutputCSV += "`n"+"CurrencyRates,Standard,REFERENCEDATE_CUR,CURRENCY1_CUR,SRCCURRENCY_CUR,RATE1_CUR"
    $OutputCSV += "`n"+"CAPTIONS,,Rates.Reference date,Rates.Destination Currency,Rates.Source Currency,Rates.Rate"
    $OutputCSV += "`n"+"TYPES,,Date,List,List,Amount"
}

# Adapted loop to display the various currency and data
foreach ($k in ($HashFxRates | Get-Member -MemberType NoteProperty).Name) {
    Write-Output "$k = $($HashFxRates.$k)"
	"Currency " + $k + " Rate against " + $BaseCCy + " on " + $FxDate + " is "+ $($HashFxRates.$k)
	# Problem of certain/uncertain quotation specific of USD 
    # if $k -eq "USD" {                                          
	#    $OutputCSV += "`n" +$CSVSep+$CSVSep + $FXDateasDate.ToString($DateFormat)+$CSVSep+$BaseCCY+$CSVSep+$k+$CSVSep+$($HashFxRates.$k)+$CSVSep
    # }
	# else {
	    $OutputCSV += "`n" +$CSVSep+$CSVSep + $FXDateasDate.ToString($DateFormat)+$CSVSep+$k+$CSVSep+$BaseCCY+$CSVSep+$($HashFxRates.$k)+$CSVSep
	# }
}

# Just save the created output into a CSV file
$OutputCSV | Out-File ./FXrate.csv

# Debug - Check Contents of CSV file
# cat ./FXRate.csv

# Final Method to import the file // The log file is the one created on the server
$cmd = """"+ $Data_Dir +"\FrontCmd.exe"" ExecWebEdgeImport /server:"""+$URL_WebSite+""" /userid:" +$Username+" /password:"+$Password+" /files:""" + $Data_Dir+"\FxRate.csv"""

# Store the command in a .bat file (Encoding ASCII guarantees that there is no odd character
$cmd | Out-File -Encoding ASCII "./go.bat"

# Execute the command
& "./go.bat"

# Delete the intermediate file
# rm go.bat

