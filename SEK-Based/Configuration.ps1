# ------------------------------------------------------------------------------
#                     Author    : eFront-Swedfund
#                     Time-stamp: "2018-08-22 14:26:28 jpdur"
# ------------------------------------------------------------------------------

# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! DEPLOYMENT SPECIFIC !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# $Data_Dir="C:\Users\jpdur\Desktop\SwedFund\FX Rate Import"     # Directory where to find the data
# $Exec_Dir="C:\Users\jpdur\Desktop\SwedFund\FX Rate Import"     # Directory where to find the exe (FrontCmd + the .ps1 scripts)
# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! DEPLOYMENT SPECIFIC !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

# Depending on the version the new header could be used
# By default we use all header format
$NewHeaderImport = 0

# To Get Encrypted Password use FrontCmd Encrypt "mypassord"
# ------------------------------------------------------------
# Result is similar to....
# Encrypted text: <enc:XUo8nBW5J7LSESrtGA9T0p3hYUZVx4ynqXMfnSUY7Lk=>

# Separator to be used for csv files /// Depends on the country specific setup
$CSVSep = ","

# Date format so that it can be read correctly into XL
$DateFormat = "dd/MM/yyyy"      #Format to be used for European Date
#$DateFormat = "MM/dd/yyyy"     #Fornat to be used for US configuration (Works for Arcano ????)

# Debug Purposes
# Write-Host "Data Dir:"$Data_Dir
# Write-Host "Exec Dir:"$Exec_Dir

# Deployment Parameters JPD Test Envt ==> NewHeader to be used
$Username        = "JPDUR"
$Password        = "enc:XUo8nBW5J7LSESrtGA9T0p3hYUZVx4ynqXMfnSUY7Lk="
$URL_WebSite     = "mandact"

# Deployment Parameters uktraining
# $Username       = "JDURAN"
# $Password       = "enc:XUo8nBW5J7LSESrtGA9T0p3hYUZVx4ynqXMfnSUY7Lk="
# $URL_WebSite    = "uktraining.frontsrv.com"

# # SwedFund Envt to be used
# $Username        = "JPDUR"
# $Password        = "enc:XUo8nBW5J7LSESrtGA9T0p3hYUZVx4ynqXMfnSUY7Lk="
# $URL_WebSite     = "https://efront.swedfund.se"
