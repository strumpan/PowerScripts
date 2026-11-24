# PowerScripts
Useful PowerScripts
- SearchFileContent.ps1

	usage: SearchFileContent.ps1 -SearchText "blabla"
	
	Searches recersivly for text, Excel, Word and PowerPoit files in the folder the skript is started from 
	
- CompareExcel.ps1

	usage: .\CompareExcel.ps1 -File1 Old.xlsx -File2 New.xlsx -KeyColumn Id
	
	Compares two Excel files using a unique key column.
    Column order does not affect the comparision.
    Results are exported to a CSV file.