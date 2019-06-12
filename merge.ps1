Class ExcelWorkBooks
{
    
    [__ComObject]$application;
    [__ComObject]$workbook;

    # error (string for format on Write-Error)
    hidden [void] error ([string]$text, [switch]$saveDoc) 
    {
        [string]$line = '-' * $text.length; 
        
        Write-Warning "$line `n$text `n$line";
        $this.Quit($saveDoc);
        exit 1;
    }

    # close Document and Quit Application.
    [void] Quit($save) 
    {
        if ($this.workbook -ne $null) 
        {
            $this.workbook.Close($save);
        }
        
        $this.application.Quit();
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.application);
    }

    # Add WorkBook .
    [object] Add([string]$path, [switch]$Exists) 
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
        
        $this.workbook   = $WB;

        return $WB;
        
    }

    # if path exists open WB
    hidden [object] OpenWorkBookIfExists([string]$path) 
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
    hidden [object] createExcelApplication ([switch]$visible)
    {
        [__ComObject]$excel = New-Object -ComObject Excel.Application;
        $excel.visible = $visible;

        return $excel;
    }

    # __init__ ([path of excel file], [create file if not exist], [made application visible or not])
    hidden ExcelWorkBooks ([string]$path, [switch]$create, [switch]$visible) 
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

#$test = [ExcelWorkBooks]::new(".\test0.csv", $false, $false);
#$test.Quit($false)
