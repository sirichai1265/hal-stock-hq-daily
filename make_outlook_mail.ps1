# Opens an Outlook draft (NOT sent) addressed to -To, attaches the Stock Daily workbook and
# pastes its three sheets into the body under the headings
#   DAILY CONTAINER STOCK / RF SEASONAL / NEW FORMAT
# Needs desktop Excel + classic Outlook. Usage:
#   powershell -ExecutionPolicy Bypass -File make_outlook_mail.ps1 -Xlsx "10-22\Stock Daily 2026-10-22.xlsx" -To someone@example.com
param(
    [Parameter(Mandatory = $true)][string]$Xlsx,
    [Parameter(Mandatory = $true)][string]$To,
    [string]$Subject = ""
)
$ErrorActionPreference = "Stop"
$full = (Resolve-Path $Xlsx).Path
if (-not $Subject) { $Subject = [IO.Path]::GetFileNameWithoutExtension($full) }
$dateText = ([regex]::Match($full, '\d{4}-\d{2}-\d{2}')).Value

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($full, 0, $true)   # read-only: never touch the report file
    $xl.CalculateFullRebuild()

    $ol = New-Object -ComObject Outlook.Application
    $mail = $ol.CreateItem(0)
    $mail.To = $To
    $mail.Subject = $Subject
    $mail.Attachments.Add($full) | Out-Null
    $mail.Display()
    $doc = $mail.GetInspector.WordEditor
    $sel = $doc.Application.Selection
    $sel.HomeKey(6) | Out-Null                  # wdStory: top of the body

    $sel.TypeText("Dear Prachanart,")
    $sel.TypeParagraph()
    $sel.TypeText("Please find attached the Container Stock Daily report as of $dateText.")
    $sel.TypeParagraph()

    $sections = @(
        @{ Title = "1. DAILY CONTAINER STOCK"; Sheet = "Daily" },
        @{ Title = "2. RF SEASONAL";           Sheet = "RF SEASONAL" },
        @{ Title = "3. NEW FORMAT";            Sheet = "NEW FORMAT" }
    )
    foreach ($s in $sections) {
        $sel.TypeParagraph()
        $sel.Font.Bold = 1
        $sel.TypeText($s.Title)
        $sel.Font.Bold = 0
        $sel.TypeParagraph()
        Write-Host "copy $($s.Sheet)"
        $ws = $wb.Worksheets.Item($s.Sheet)
        $ws.Activate()
        $ws.UsedRange.Copy() | Out-Null
        Start-Sleep -Milliseconds 800
        Write-Host "paste $($s.Sheet)"
        $sel = $doc.Application.Selection
        $sel.Paste()                            # selection ends up just after the pasted table
        Start-Sleep -Milliseconds 800
        $sel.TypeParagraph()
    }
    "Draft opened in Outlook: To=$To  Subject=$Subject  (not sent)"
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
