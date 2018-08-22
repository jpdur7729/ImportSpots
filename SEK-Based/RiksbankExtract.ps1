# ------------------------------------------------------------------------------
#                     Author    : eFront-SwedFund
#                     Time-stamp: "2018-08-22 15:17:15 jpdur"
# ------------------------------------------------------------------------------
                                                 
# ---------------------------------------------------------------------------
# The received parameter is the directory where the script is to be executed
# if nothing is received then let's use the directory where the script is
# found .... as a default
# ---------------------------------------------------------------------------
param(
	  [Parameter(Mandatory=$false)] [string] $Exec_Dir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
)

Function Release-Ref ($ref) 
    {
        ([System.Runtime.InteropServices.Marshal]::ReleaseComObject(
        [System.__ComObject]$ref) -gt 0)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers() 
    }

Function ConvertCSV-ToExcel
{
<#
  .SYNOPSIS
    Converts one or more CSV files into an excel file.

  .DESCRIPTION
    Converts one or more CSV files into an excel file. Each CSV file is imported into its own worksheet with the name of the
    file being the name of the worksheet.

  .PARAMETER inputfile
    Name of the CSV file being converted

  .PARAMETER output
    Name of the converted excel file

  .EXAMPLE
  Get-ChildItem *.csv | ConvertCSV-ToExcel -output 'report.xlsx'

  .EXAMPLE
  ConvertCSV-ToExcel -inputfile 'file.csv' -output 'report.xlsx'

  .EXAMPLE
  ConvertCSV-ToExcel -inputfile @("test1.csv","test2.csv") -output 'report.xlsx'

  .NOTES
  Author: Boe Prox
  Date Created: 01SEPT210
  Last Modified:

#>

#Requires -version 2.0
[CmdletBinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = 'low',
	DefaultParameterSetName = 'file'
    )]
Param (
    [Parameter(
     ValueFromPipeline=$True,
     Position=0,
     Mandatory=$True,
     HelpMessage="Name of CSV/s to import")]
     [ValidateNotNullOrEmpty()]
    [array]$inputfile,
    [Parameter(
     ValueFromPipeline=$False,
     Position=1,
     Mandatory=$True,
     HelpMessage="Name of excel file output")]
     [ValidateNotNullOrEmpty()]
    [string]$output    
    )

Begin {     
    #Configure regular expression to match full path of each file
    [regex]$regex = "^\w\:\\"
    
    #Find the number of CSVs being imported
    $count = ($inputfile.count -1)
   
    #Create Excel Com Object
    $excel = new-object -com excel.application
    
    #Disable alerts
    $excel.DisplayAlerts = $False

    #Show Excel application
    $excel.Visible = $False

    #Add workbook
    $workbook = $excel.workbooks.Add()

    #Remove other worksheets
    $workbook.worksheets.Item(2).delete()
    #After the first worksheet is removed,the next one takes its place
    $workbook.worksheets.Item(2).delete()   

    #Define initial worksheet number
    $i = 1
    }

Process {
    ForEach ($input in $inputfile) {
        #If more than one file, create another worksheet for each file
        If ($i -gt 1) {
            $workbook.worksheets.Add() | Out-Null
            }
        #Use the first worksheet in the workbook (also the newest created worksheet is always 1)
        $worksheet = $workbook.worksheets.Item(1)
        #Add name of CSV as worksheet name
        $worksheet.name = "$((GCI $input).basename)"

        #Open the CSV file in Excel, must be converted into complete path if no already done
        If ($regex.ismatch($input)) {
            $tempcsv = $excel.Workbooks.Open($input) 
            }
        ElseIf ($regex.ismatch("$($input.fullname)")) {
            $tempcsv = $excel.Workbooks.Open("$($input.fullname)") 
            }    
        Else {    
            $tempcsv = $excel.Workbooks.Open("$($pwd)\$input")      
            }
        $tempsheet = $tempcsv.Worksheets.Item(1)
        #Copy contents of the CSV file
        $tempSheet.UsedRange.Copy() | Out-Null
        #Paste contents of CSV into existing workbook
        $worksheet.Paste()

        #Close temp workbook
        $tempcsv.close()

        #Select all used cells
        $range = $worksheet.UsedRange

        #Autofit the columns
        $range.EntireColumn.Autofit() | out-null
        $i++
        } 
    }        

End {
    #Save spreadsheet
    $workbook.saveas("$pwd\$output")

    Write-Host -Fore Green "File saved to $pwd\$output"

    #Close Excel
    $excel.quit()  

    #Release processes for Excel
    $a = Release-Ref($range)
    }
}

# -----------------------------------------------------------------------------------------------------
# Move to the desired Directory i.e the directory where FrontCmd.exe and config files can be found
# This is based on the parameter received
# -----------------------------------------------------------------------------------------------------
cd ($Exec_Dir)                                   

# Configuration file to get all the corresponding parameters
. ./Configuration.ps1

# ------------------------ Creation of the Extract Parameter SOAP msg file  ----------------------------------
$fileName = $Exec_Dir + "\SwedenSoap3.xml"  
$xmlDoc = [System.Xml.XmlDocument](Get-Content $fileName);
# $xmlDoc.Envelope.Body.getCrossRates.crossRequestParameters.datefrom
# $xmlDoc.Envelope.Body.getCrossRates.crossRequestParameters.dateto
write-host "Data before Processing " $xmlDoc.Envelope.Body.getCrossRates.crossRequestParameters.dateto
                          
# Modify the dates in the previous generation document Today and 4 physical days before
#$xmlDoc.Envelope.Body.getCrossRates.crossRequestParameters.datefrom=((Get-Date).AddDays(-4).ToString("yyyy-MM-dd"))
# $xmlDoc.Envelope.Body.getCrossRates.crossRequestParameters.datefrom= (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")
$xmlDoc.Envelope.Body.getCrossRates.crossRequestParameters.datefrom= (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")
$xmlDoc.Envelope.Body.getCrossRates.crossRequestParameters.dateto  = (Get-Date).ToString("yyyy-MM-dd")
                          
write-host "Data after Processing " $xmlDoc.Envelope.Body.getCrossRates.crossRequestParameters.dateto
                          
# # Due to testing before 9:30
# $xmlDoc.Envelope.Body.getCrossRates.crossRequestParameters.datefrom= ((Get-Date).AddDays(-1).ToString("yyyy-MM-dd"))
# $xmlDoc.Envelope.Body.getCrossRates.crossRequestParameters.dateto  = ((Get-Date).AddDays(-1).ToString("yyyy-MM-dd"))
                          
# Modify the updated document
$xmlDoc.save($fileName)

# ------------------------ Extract Data from Riksbank ----------------------------------
# Call of the extract from Riksbank - Example with hardcoded soap request file
# $cmd ="curl --header ""Content-Type: text/xml;charset=UTF-8"" --header ""SOAPAction:getCrossRates"" --data @SwedenSoap3.msg https://swea.riksbank.se:443/sweaWS/services/SweaWebServiceHttpSoap12Endpoint -k -o ResultExtract.xml"
$cmd ="curl --header ""Content-Type: text/xml;charset=UTF-8"" --header ""SOAPAction:getCrossRates"" --data @"+$fileName+" https://swea.riksbank.se:443/sweaWS/services/SweaWebServiceHttpSoap12Endpoint -k -o ResultExtract.xml"
                                                                                                      
# Store the command in a .bat file (Encoding ASCII guarantees that there is no odd character          
$cmd | Out-File -Encoding ASCII "./go.bat"                                                            
                                                                                                      
# Execute the command                                                                                 
& "./go.bat"                                                                                          
                                                                                                      
# Delete the intermediate file                                                                        
rm go.bat                                                                                             
                                                                                                      
# ------------------------ Preapre the value of the EUR/SEK FX Rate  ----------------------------------
$SEKEURRate       = 0.0                                                                               
$RiksbankCCYList  = New-Object System.Collections.ArrayList # List of Ccy provided by Riksbank        
$Today            = (Get-Date).ToString("yyyy-MM-dd")                                                                                  
                                                     
# ------------------------ Preparing the csv Import file  ----------------------------------
# SilentlyContinue does not give any error message is the file is not found...
rm  ./FXrate.csv -ErrorAction SilentlyContinue       
                                                     
# 1st line import CSV file                           
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
                                                     
# We normalise the header with the correct separator 
$OutputCSV = $OutputCSV -replace ",",$CSVSep         
                                                     
# ------------------------ Processing Data from Riksbank ----------------------------------
#Read from file as generated by Riksbank             
[xml]$z = Get-Content ".\ResultExtract.xml"          
                                                     
# Loop to extract the data                           
foreach( $PerCCy in $z.Envelope.Body.getCrossRatesResponse.return.groups.series)
{                                                    
    # write-host "test" + $PerCCy.seriesname         
	# Process the Seriesname                         
    $BaseCCY = $PerCCy.seriesname.Substring(2,3)
    $CCY     = $PerCCy.seriesname.Substring(10,3)
	If (($BaseCCY -eq "SEK") -And ($CCY.trim().length -eq 3) )
    {
		foreach ($DateRate in $PerCCy.resultrows) {

            # Debug - Check contents of date
			# Write-Host "/" $DateRate.date "/" $DateRate.date.length

	 		# Change the string to date in order to be able to later choose the right format
			# The $DateRate.date has actually more than the required 10 characters ==> Hence the checks
			$DateFXRate  = [datetime]::ParseExact($DateRate.date.Substring(0,10), "yyyy-MM-dd", [CultureInfo]::InvariantCulture)

            # $ValueFXRate = $DateRate.value.trim()
            $ValueFXRate = $DateRate.value

			Write-Host "Verif" $Today $DateRate.date $ValueFXRate

			if ( $Today -eq $DateRate.date) {

				# Add the currency in the list of currencies provided by Riks Bank
				[void] $RiksbankCCYList.Add($CCY)

				# Store the EUR/SEK FX Rate
				if ($CCY -eq "EUR") {
					$SEKEURRate = [convert]::ToDecimal($ValueFXRate)
				}

				# SwedFund uses the , as the decimal separator so we replace the . (\. to prevent regexpreson)
				# $ValueFXRate = $ValueFXRate -replace "\.",","

				# Debug the data to be written in the csv file
				Write-Host "Verification" $BaseCCY $CCY $DateFXRate.ToString($DateFormat) $ValueFXRate

	 			$OutputCSV += "`n" +$CSVSep+$CSVSep + $DateFXRate.ToString($DateFormat)+$CSVSep+$CCY+$CSVSep+$BaseCCY+$CSVSep+$ValueFXRate+$CSVSep
			}
		}
    }
}

# ----------------------------------------------------------------------------------------------------
# ----------------------------------- Extract the Data from data.fixer.io ----------------------------
# ----------------------------------------------------------------------------------------------------
# We built the request by default for today's rate and all available symbols
# $request = "2017-12-31?access_key=eca17521f4e211d09ab357c6cd9585dc&base=EUR&symbols=USD,CAD,EUR,GBP"
  $request = (Get-Date -UFormat "%Y-%m-%d") + "?access_key=eca17521f4e211d09ab357c6cd9585dc&base=EUR" # JPD Key
# $request = (Get-Date -UFormat "%Y-%m-%d") + "?access_key=9038512b080e95aba278254a579bbe16&base=EUR&symbols=USD,CAD,EUR,GBP" # Niclas key

# ~~~~~~~~~~~~~~~~~~~~~~ Start Extract ~~~~~~~~~~~~~~~~~~~~~~~~~~
# Extract the FX rates from the data.fixer.io Web Site which generated json
$extractcmd = "wget ""http://data.fixer.io/api/"+ $request + """ -o LogFile.txt"
                                                                    
# We delete the files that we will use for the extract                                                                                                                    
rm Result.json -ErrorAction SilentlyContinue                                                                                                                              
                                                                                                                                                                          
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
                                                                                                                                                                          
# -------------------------------------------------------------------------------------------------------------                                                           
# Try to process the data which is not exactly an HashTable but something close as per the article below                                                                  
# https://stackoverflow.com/questions/22002748/hashtables-from-convertfrom-json-have-different-type-from-powershells-built-in-h                                           
# The contents is adapated to the case                                                                                                                                    
# -------------------------------------------------------------------------------------------------------------                                                           
# # $HashFxRates.root | select * | ft -AutoSize # Does mot work                                                                                                           
# $HashFxRates | select * | ft -AutoSize        # Works                                                                                                                   
                                                                                                                                                                          
# SilentlyContinue does not give any error message is the file is not found...                                                                                            
rm  ./FXrate.csv -ErrorAction SilentlyContinue                                                                                                                            
                                                                                                                                                                          
# Adapted loop to display the various currency and data                                                                                                                   
foreach ($k in ($HashFxRates | Get-Member -MemberType NoteProperty).Name) {                                                                                               
    Write-Output "$k = $($HashFxRates.$k)"                                                                                                                                
	# "Currency " + $k + " Rate against " + $BaseCCy + " on " + $FxDate + " is "+ $($HashFxRates.$k)                                                                      
	# We only add the currency if it has not already been provided by RiksBank                                                                                            
	if ( -not ($RiksbankCCYList | Where-Object { $k -like $_} )) {                                                                                                        
	    $OutputCSV += "`n" +$CSVSep+$CSVSep + $FXDateasDate.ToString($DateFormat)+$CSVSep+$k+$CSVSep+"SEK"+$CSVSep+[math]::Round(($HashFxRates.$k)*$SEKEURRate,4)+$CSVSep 
	}
}

# ---------------------------- Package all the data into the CSV File -------------------------

# Just save the created output into a CSV file
$OutputCSV | Out-File -Encoding UTF8 ./FXrate.csv

# Convert to XLS in order to sort all the conversion issues
# ConvertCSV-ToExcel "FXRate.csv" "FXRate.xlsx"

# ------------------------ Import Data into eFront  ----------------------------------
# Final Method to import the file // The log file is the one created on the server
# $cmd = """"+ $Exec_Dir +"\FrontCmd.exe"" ExecWebEdgeImport /server:"""+$URL_WebSite+""" /userid:" +$Username+" /password:"+$Password+" /files:""" + $Exec_Dir+"\FxRate.xlsx"""
$cmd = """"+ $Exec_Dir +"\FrontCmd.exe"" ExecWebEdgeImport /server:"""+$URL_WebSite+""" /userid:" +$Username+" /password:"+$Password+" /files:""" + $Exec_Dir+"\FxRate.csv"""

# Store the command in a .bat file (Encoding ASCII guarantees that there is no odd character
$cmd | Out-File -Encoding ASCII "./go.bat"

# Execute the command
& "./go.bat"

# Delete the intermediate file
rm go.bat

# Cleanup the log files
mv *.log log

# Debug Infop ==> List of RiksBank CCy
# $RiksbankCCYList
