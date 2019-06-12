# ExcelWorkBooks
> ## ExcelWorkBooks class for easy use Excel WorkBook in PowerShell.
### If you need make some script use excel you can fast create com object with work book and application.
#### example (if we need create new Book):
```PowerShell

# __init__ ([path of excel file], [create file if not exist], [made application visible or not])

[ExcelWorkBooks]$WB = [ExcelWorkBooks]::new(".\Foo.csv", $true, $false);
$WB.Application.Range("A1") = "Hello PS!";
$WB.WorkBook.ActiveSheet.Cells(1,2) = 0;

# Kill Com objects(Boolean[Save doc or not])
$WB.Quit($true)

```
