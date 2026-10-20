$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$files = @(
    "c:\Antigravity_Thanh\San_Pham\De_Kiem_Tra_Hoa_12_Ma638_HocSinh.docx",
    "c:\Antigravity_Thanh\San_Pham\De_Kiem_Tra_Hoa_12_Ma638_GiaoVien.docx",
    "c:\Antigravity_Thanh\San_Pham\De_Kiem_Tra_Hoa_12_Ma638_DayTrenManHinh.docx"
)

foreach ($f in $files) {
    if (Test-Path $f) {
        $pdfPath = [System.IO.Path]::ChangeExtension($f, ".pdf")
        Write-Host "Converting $f to $pdfPath..."
        $doc = $word.Documents.Open($f)
        $doc.SaveAs([ref]$pdfPath, [ref]17)
        $doc.Close([ref]$false)
        Write-Host "Success: $pdfPath"
    } else {
        Write-Host "File not found: $f"
    }
}

$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Host "Conversion finished."
