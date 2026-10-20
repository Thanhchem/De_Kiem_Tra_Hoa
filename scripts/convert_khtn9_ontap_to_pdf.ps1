$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$files = @(
    "c:\Antigravity_Thanh\San_Pham\De_Kiem_Tra_KHTN_9_Giua_HK1_HocSinh.docx",
    "c:\Antigravity_Thanh\San_Pham\De_Kiem_Tra_KHTN_9_Giua_HK1_GiaoVien.docx",
    "c:\Antigravity_Thanh\San_Pham\De_Kiem_Tra_KHTN_9_Giua_HK1_DayTrenManHinh.docx"
)

foreach ($f in $files) {
    if (Test-Path $f) {
        $pdfPath = [System.IO.Path]::ChangeExtension($f, ".pdf")
        Write-Host "Converting $f to $pdfPath..."
        $doc = $word.Documents.Open($f)
        $doc.SaveAs([ref]$pdfPath, [ref]17)
        $doc.Close([ref]$false)
        Write-Host "Success: $pdfPath"
        
        # Đồng bộ PDF sang OneDrive
        $od1 = "C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9\KHTN9-DEKIEMTRADINHKI"
        $od2 = "C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9"
        if (Test-Path $od1) {
            Copy-Item -Path $pdfPath -Destination $od1 -Force
            Write-Host "Copied PDF to $od1"
        }
        if (Test-Path $od2) {
            Copy-Item -Path $pdfPath -Destination $od2 -Force
            Write-Host "Copied PDF to $od2"
        }
    } else {
        Write-Host "File not found: $f"
    }
}

$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Host "Conversion finished."
