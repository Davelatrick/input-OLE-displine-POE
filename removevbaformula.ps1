function Convert-ExcelFormulasToValues {
    param (
        [string]$filePath
    )

    # Create Excel.Application object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false  # Suppress alerts

    # Create Shell.Application object to unblock the file
    $shell = New-Object -ComObject Shell.Application
    $folder = Split-Path $filePath
    $shellfolder = $shell.Namespace($folder)
    $fileObj = $shellfolder.ParseName((Split-Path $filePath -Leaf))

    # Check if the file is blocked
    if ($fileObj.ExtendedProperty("{098F2470-BAE0-11CD-B579-08002B30BFEB} 2") -eq 1) {
        # Unblock the file
        $fileObj.InvokeVerb("Unblock")
        Write-Output "File has been successfully unblocked: $filePath"
    } else {
        Write-Output "File is not blocked or has been unblocked: $filePath"
    }

    # Release Shell.Application object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
    Remove-Variable shell

    # Create new file path for the cleaned version
    $newFilePath = [System.IO.Path]::Combine(
        [System.IO.Path]::GetDirectoryName($filePath),
        [System.IO.Path]::GetFileNameWithoutExtension($filePath) + "_cleaned.xlsx"
    )

    # Open workbook
    $workbook = $excel.Workbooks.Open($filePath)

    # Remove all VBA code
    try {
        if ($workbook.HasVBProject) {
            $workbook.VBProject.VBComponents | ForEach-Object {
                $workbook.VBProject.VBComponents.Remove($_)
            }
            Write-Output "VBA code has been removed"
        } else {
            Write-Output "No VBA code found in the workbook"
        }
    }
    catch {
        Write-Output "Could not remove VBA code. Trust access to the VBA project object model might not be enabled."
    }

    foreach ($worksheet in $workbook.Worksheets) {
        $usedRange = $worksheet.UsedRange

        # Store information about merged cells
        $mergedRanges = @()
        foreach ($range in $worksheet.UsedRange.MergeCells) {
            $mergedRanges += [PSCustomObject]@{
                Address = if ($range.Address) { $range.Address(0,0,1) } else { "NoAddress" }
                TopLeftValue = if ($range.Cells(1,1).Value2) { $range.Cells(1,1).Value2 } else { "NoValue" }
            }
        }

        # Copy values and paste
        $usedRange.Copy()
        $usedRange.PasteSpecial(-4163) # xlPasteValues

        # Re-merge cells
        foreach ($mergedRange in $mergedRanges) {
            $rangeToMerge = $worksheet.Range($mergedRange.Address)
            if ($rangeToMerge.Cells(1,1).Value2 -ne $mergedRange.TopLeftValue) {
                $rangeToMerge.Cells(1,1).Value2 = $mergedRange.TopLeftValue
            }
            $rangeToMerge.Merge()
        }

        # Paste formats
        $usedRange.Copy()
        $usedRange.PasteSpecial(-4122) # xlPasteFormats

        # Clear clipboard
        $excel.CutCopyMode = [Microsoft.Office.Interop.Excel.XlCutCopyMode]::xlCopy
    }

    # Save as xlsx (this format doesn't support macros)
    $workbook.SaveAs($newFilePath, 51)  # 51 = xlsx format
    $workbook.Close()

    # Quit Excel
    $excel.Quit()

    # Release Excel.Application object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable excel

    Write-Output "Process completed. Cleaned file saved as: $newFilePath"
}

# User input file path
$filePath = Read-Host "Please enter the full path of the Excel file"

# Call function
Convert-ExcelFormulasToValues -filePath $filePath