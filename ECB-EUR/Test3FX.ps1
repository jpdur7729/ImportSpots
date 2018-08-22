# ------------------------------------------------------------------------------
#                     Author    : eFront-Mastek
#                     Time-stamp: "2018-06-25 10:21:52 jpdur"
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

# wget http://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml -P %workingdir%
# Extract the FX rates from the ECB web
$extractcmd = "wget http://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml -P """ + $Exec_Dir + """"

# As we are working om the ($Exec_Dir) we want to be sure that the output file does not exist
rm eurofxref-daily.xml

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
            
# Check the required file is available
# cat eurofxref-daily.xml
              
$filter = "eurofxref-daily*.xml"
              
# Identify the file(s) in the Process Server and Process it accordingly
# In that specific case it happens only if there is a TmpDvlperSalesMIS.xlsx file
Get-ChildItem -path $Exec_Dir -filter $filter | Foreach-Object {

    #Destination file with extension
	$Filename_Process     = $_.Name
	$Filename_NoExtension = $_.BaseName

	# Read File Contents
	$Result = Get-Content ($Data_Dir+"\"+$Filename_Process)

    # Prepare the result buffer with the header of the result
	# $OutputCSV = "Reference date" + $CSVSep + "Source Currency" + $CSVSep + "Destination Currency" + $CSVSep + "Rate" + $CSVSep + "Description"
	# String is single quoted to prevent interpreation of $ (cf. http://www.rlmueller.net/PowerShellEscape.htm)
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
                    
	# Extract the lines of the contents
	$ContentsReached = $false ;
                    
	 for($i=0;$i-le $Result.length-1;$i++){
	 	# Skip the 1st lines of the header up to the point where we have found contents
	 	if (-not $ContentsReached) {
	 		$ContentsReached = ($Result[$i] -like "*Cube time='*")
	 		# If found then we extract the date accordingly which will be used thereafter
	 		if ($ContentsReached) {
                    
	 			# We identify the line                                                                         
	 			$Line = $Result[$i]                                                                            
                                                                                                               
	 			# We extract te date by searching for the expression "Cube time='"                             
	 			$ExprDate = "Cube time='"                                                                      
	 			$FXDateString = $Line.substring($Line.IndexOf($ExprDate)+$ExprDate.Length,10)
                    
	 			# Change the string to date in order to be able to later choose the right format
	 			$FxDateasDate = [datetime]::ParseExact($FXDateString,"yyyy-MM-dd", $null)
	 		}       
	 	}           
	 	else {      
	 		# Let's process the current line if it contents <Cube currency
	 		$Line = $Result[$i]
                    
	 		# Debug: Verify the contents
	 		# $Line 
                    
	 		if ($Line -like "*<Cube currency=*" ){
                    
	 			# Extract the currency 1st by searching the expression currency='
                $ExprCCY = "<Cube currency='"
                $CCY = $Line.SubString($Line.IndexOf($ExprCCY)+$ExprCCy.Length,3)
                    
	 		    # Extract the rate by searching the expression rate='
	 		    $ExprRate = "rate='"
                $StartRate = $Line.IndexOf($ExprRate)+$ExprRate.Length
	 		    $Rate = $Line.Substring($StartRate,$Line.LastIndexOf("'") - $StartRate)
                    
	 		    # We add the corresponding String to the results
	 			# Adding cr to a powershell line https://blogs.technet.microsoft.com/heyscriptingguy/2014/09/07/powertip-new-lines-with-powershell/
				# We jump the 1st 2 columns when adding the data                             
	 			# OLD VERSION $OutputCSV += "`n" +$CSVSep+$CSVSep + $FXDateasDate.ToString($DateFormat)+$CSVSep+"EUR"+$CSVSep+$CCY+$CSVSep+$Rate+$CSVSep
				# By convention EUR = SRCCURRENCY V
	 			$OutputCSV += "`n" +$CSVSep+$CSVSep + $FXDateasDate.ToString($DateFormat)+$CSVSep+$CCY+$CSVSep+"EUR"+$CSVSep+$Rate+$CSVSep
	 		}
	 		# We skip the lines which have no contents
	 	    # else { "End File ...." }
	 	}
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
	rm go.bat

	# Move the log file to the corresponding directory
    mv -force *.log ../Log

	# Prepare the steps by which the extracted data will be kept with the corresponding date
	$Filename_NoExtension = $_.BaseName
	mv -force $FileName_Process ("../Data/"+$Filename_NoExtension+$FxDateasDate.ToString("yyyyMMdd")+".xml")

} # End full loop on the list of a given statement

