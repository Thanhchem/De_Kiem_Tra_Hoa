$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$files = @(
    "c:\Antigravity_Thanh\San_Pham\De_Cuong_KHTN_9_Giua_HK1_Hoa_Ly_HocSinh.docx",
    "c:\Antigravity_Thanh\San_Pham\De_Cuong_KHTN_9_Giua_HK1_Hoa_Ly_GiaoVien.docx",
    "c:\Antigravity_Thanh\San_Pham\De_Cuong_KHTN_9_Giua_HK1_Hoa_Ly_DayTrenManHinh.docx"
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
        
        # Copy to OneDrive
        $odList = @(
            "c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9\",
            "c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9\KHTN9-DEKIEMTRADINHKI\"
        )
        foreach ($od in $odList) {
            if (Test-Path $od) {
                $odFile = Join-Path $od ([System.IO.Path]::GetFileName($pdfPath))
                Copy-Item $pdfPath $odFile -Force
                Write-Host "Copied to OneDrive: $odFile"
            }
        }
    } else {
        Write-Host "File not found: $f"
    }
}

$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Host "PDF Conversion for KHTN 9 (Hoa & Ly) finished successfully."
