# PowerScripts
Useful PowerScripts
- SearchFileContent.ps1

	usage: SearchFileContent.ps1 -SearchText "blabla"
	
	Searches recersivly for text, Excel, Word and PowerPoit files in the folder the skript is started from 
	
- Compare-Excel.ps1

	usage: .\Compare-Excel.ps1 -File1 Old.xlsx -File2 New.xlsx
	
	Compares rows by Excel row number, including gaps for empty rows.
	The CSV identifies differences using the worksheet row number. Inserting
	or deleting a row shifts the subsequent row-by-row comparisons.

	Use `-KeyColumn Id` to compare records by a unique key instead.
	Column order does not affect the comparison. Results are exported to a CSV file.