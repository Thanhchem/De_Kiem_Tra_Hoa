$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$files = @(
    "c:\Antigravity_Thanh\San_Pham\De_To36_Hoa12_PinDien_DienPhan_HocSinh.docx",
    "c:\Antigravity_Thanh\San_Pham\De_To36_Hoa12_PinDien_DienPhan_GiaoVien.docx",
    "c:\Antigravity_Thanh\San_Pham\De_To36_Hoa12_PinDien_DienPhan_DayTrenManHinh.docx"
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
            "c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\12\",
            "c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\12\kiem-tra-dinh-ky\"
        )
        foreach ($od in $odList) {
            if (Test-Path $od) {
                $odFile = Join-Path $od ([System.IO.Path]::GetFileName($f))
                Copy-Item $f $odFile -Force
                Write-Host "Copied DOCX to OneDrive: $odFile"

                $odPdf = Join-Path $od ([System.IO.Path]::GetFileName($pdfPath))
                Copy-Item $pdfPath $odPdf -Force
                Write-Host "Copied PDF to OneDrive: $odPdf"
            }
        }
    } else {
        Write-Host "File not found: $f"
    }
}

$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Host "PDF Conversion for To 36 Hoa 12 finished successfully."
