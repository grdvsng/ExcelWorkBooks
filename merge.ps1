Class ExcelWorkBooks
{
    
    [array]$workbooks = {}; #AllWB
    [__ComObject]$application;
    [__ComObject]$workbook;

    # error (string for format on Write-Error)
    [void] error ([string]$text, [boolean]$saveDoc) 
    {
        [string]$line = '-' * $text.length; 
        
        Write-Host "$line `n$text `n$line";
        $this.Quit($saveDoc);
        exit 1;
    }

    # close Document and Quit Application.
    [void] Quit($Save) 
    {
        if ($this.workbook -ne $null) 
        {
            for ($n=1; $n -ne $this.workbooks.Count; $n++) 
            {
                
                [__ComObject]$WB = $this.workbooks[$n];
                $WB.Close();
            }
        }
        
        $this.application.Quit();
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.application);
    }

    # Add WorkBook .
    [object] Add([string]$path, [boolean]$Exists) 
    {
        if ($Exists) 
        {
            [__ComObject]$WB = $this.OpenWorkBookIfExists($path);
        }
        else
        {
            [__ComObject]$WB = $this.application.WorkBooks.Add();
            $WB.SaveAs($path);
        }
        
        $this.workbooks += $WB;
        $this.workbook   = $WB;

        return $WB;
        
    }

    # if path exists open WB
    [object] OpenWorkBookIfExists([string]$path) 
    {        
        $exists = Get-Childitem $path;

        if ($exists -eq $null) 
        {
            return $this.error("file: '$path', not found.", $false);
        }
        else   
        {
            return $this.application.WorkBooks.Open($path);
        }
    }

    # Comobj create
    [object] createExcelApplication ([boolean]$visible)
    {
        [__ComObject]$excel = New-Object -ComObject Excel.Application;
        $excel.visible = $visible;

        return $excel;
    }

    # __init__ ([path of excel file], [create file if not exist], [made application visible or not])
    ExcelWorkBooks ([string]$path, [boolean]$create, [boolean]$visible) 
    {
        $this.application = $this.createExcelApplication($visible);
        
        if ($create) 
        {
            $this.workbook = $this.Add($path, $false);
        }
        else
        {
            $this.workbook = $this.Add($path, $true);
        }
    }
}

#$test = [ExcelWorkBooks]::new("C:\projects\forJob\dataMerge\test0.csv", $false, $false);
#$test.Add("C:\projects\forJob\dataMerge\test1.csv", $true)
#$test.Quit($false)