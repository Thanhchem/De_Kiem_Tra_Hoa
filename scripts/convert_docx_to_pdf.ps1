$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$files = @(
    "c:\Antigravity_Thanh\San_Pham\De_Cuong_KHTN_6_On_Tap_GK1_HocSinh.docx",
    "c:\Antigravity_Thanh\San_Pham\De_Cuong_KHTN_6_On_Tap_GK1_GiaoVien.docx",
    "c:\Antigravity_Thanh\San_Pham\De_Cuong_KHTN_6_On_Tap_GK1_DayTrenManHinh.docx",
    "c:\Antigravity_Thanh\San_Pham\De_Cuong_KHTN_9_Hoa_Hoc_GK1_HocSinh.docx",
    "c:\Antigravity_Thanh\San_Pham\De_Cuong_KHTN_9_Hoa_Hoc_GK1_GiaoVien.docx",
    "c:\Antigravity_Thanh\San_Pham\De_Cuong_KHTN_9_Hoa_Hoc_GK1_DayTrenManHinh.docx"
)

foreach ($f in $files) {
    if (Test-Path $f) {
        $pdfPath = [System.IO.Path]::ChangeExtension($f, ".pdf")
        Write-Host "Converting $f to $pdfPath..."
        $doc = $word.Documents.Open($f)
        # 17 corresponds to wdFormatPDF
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
