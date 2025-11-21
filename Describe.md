A set of powershell functions that that help translating product guids to string ids and frinedly names 

Get-MicrosoftProductSheet should download a copy of the csv tot he downloads folder from this URI:
https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv

replace the file if it already exists so it also serves as a way to update the sheet to make sure the latest version is available

Import-MicrosoftProductSheet should load the CSV into memory somehow

Get-MSProduct should facilitate some way of finding a product from the sheet in memory by guid or search string perhaps

review the copy of the CSV I have for how to query the data
