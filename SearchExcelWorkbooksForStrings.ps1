#Issue! - Currently, this only returns searches from the last workbook searched recursively from the directory

#First, open the PowerShell ISE with the "Run as Administrator" command
#Then, run the following command to allow scripts to run:
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine

Function Search-Excel {
    [cmdletbinding()]
    Param (
        [parameter(Mandatory, ValueFromPipeline)]
        [ValidateScript({
            Try {
                If (Test-Path -Path $_) {$True}
                Else {Throw "$($_) is not a valid path!"}
            }
            Catch {
                Throw $_
            }
        })]
        [string]$Source,
        [parameter(Mandatory)]
        [string]$SearchText
        #You can specify wildcard characters (*, ?)
    )
    $Excel = New-Object -ComObject Excel.Application
    Try {
        $Source = Convert-Path $Source
    }
    Catch {
        Write-Warning "Unable locate full path of $($Source)"
        BREAK
    }
    $Workbook = $Excel.Workbooks.Open($Source)
    ForEach ($Worksheet in @($Workbook.Sheets)) {
        # Find Method https://msdn.microsoft.com/en-us/vba/excel-vba/articles/range-find-method-excel
        $Found = $WorkSheet.Cells.Find($SearchText) #What
        If ($Found) {
            # Address Method https://msdn.microsoft.com/en-us/vba/excel-vba/articles/range-address-property-excel
            $BeginAddress = $Found.Address(0,0,1,1)
            #Initial Found Cell
            [pscustomobject]@{
                WorkSheet = $Worksheet.Name
                Column = $Found.Column
                Row =$Found.Row
                Text = $Found.Text
                Address = $BeginAddress
            }
            Do {
                $Found = $WorkSheet.Cells.FindNext($Found)
                $Address = $Found.Address(0,0,1,1)
                If ($Address -eq $BeginAddress) {
                    BREAK
                }
                [pscustomobject]@{
                    WorkSheet = $Worksheet.Name
                    Column = $Found.Column
                    Row =$Found.Row
                    Text = $Found.Text
                    Address = $Address
                }                
            } Until ($False)
        }
        Else {
            Write-Warning "[$($WorkSheet.Name)] Nothing Found!"
        }
    }
    $workbook.close($false)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable excel -ErrorAction SilentlyContinue
}

#In the following line, adjust the -Path value and the -SearchText value
#Do not forget to add an asterisk at the end of the directory for -Path
Get-ChildItem -Path "C:\Users\Charles\Desktop\TestFolder\*" -Recurse -Include *.xls, *.xlsx, *.xlsm | Select-Object -Property Directory, Name | ForEach-Object { "{0}\{1}" -f $_.Directory, $_.Name } | Search-Excel -SearchText 23