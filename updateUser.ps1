# Initial setup
$sheetPath = "C:\Users\tom.wallis\Downloads\SetOSCEXOEmailSignature\Book1.xlsx"
$sigPath = "C:\Users\tom.wallis\TomWallis.htm"
$sheetName = "Sheet1"
$row = 2             # NOTE: ASSUMING HEADERS MAKES AN ASS OUT OF YOU AND ME.
$column = 1
# Import the Powershell functions for configuring office365 signature settings
Import-Module .\SetOSCEXOEmailSignature.psm1    


# Open the Excel worksheet for working on
$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $false    # Don't start Excel! Excel should run in the background.
$workbook = $objExcel.Workbooks.Open($sheetPath)
$worksheet = $workbook.sheets.item($sheetName)

# Iterate over the cells and set signatures as necessary. 
while ($worksheet.cells.Item($row, $column).text -ne "") { # TODO: is this condition right?

	# Extract data from the sheet
	$user = $worksheet.cells.Item($row, $column).text
	$pass = $worksheet.cells.Item($row, $column+1).text
	$securePass = convertto-securestring $pass -AsPlainText -Force
	$credentials = New-Object System.Management.Automation.PSCredential ($user, $securepass)

	# Make the connection and update the signature
	Connect-OSCEXOWebService -Credential $credentials # Create the connection to the server (and log in) 
	Set-OSCEXOEmailSignature -TextSignature $null -HtmlSignature $null -Verbose     # Clear old signature
	Set-OSCEXOEmailSignature -HtmlSignature (Get-Content $sigPath | Out-String) -Verbose -Force    # Add new signature

	# Move on to the next position in the sheet
	$row = $row+1
}
