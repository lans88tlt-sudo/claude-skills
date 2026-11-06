<#
    Шаг 2-3 конвейера скила excel-oformlenie.
    Открывает .xlsx, подгоняет высоту НЕобъединённых строк силами самого Excel,
    настраивает печать и выгружает PDF для визуальной проверки.

    ВАЖНО: объединённые ячейки пропускаются намеренно — AutoFit на них не работает
    и схлопывает строку. Их высота должна быть задана заранее в Python.

    Файл обязан лежать в UTF-8 **с BOM**, иначе PowerShell 5.1 прочтёт кириллицу как ANSI.

    Пример:
      powershell -File autofit_and_pdf.ps1 -Path "D:\...\Таблица.xlsx" -Pdf "D:\tmp\out.pdf"
#>
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $false)][string]$Pdf,
    [Parameter(Mandatory = $false)][string]$TitleRow = '',      # напр. '$5:$5' — повтор шапки при печати
    [Parameter(Mandatory = $false)][switch]$Portrait,
    [Parameter(Mandatory = $false)][int]$MinRowHeight = 16,
    [Parameter(Mandatory = $false)][int]$FitTall = 0        # >0 — ужать лист и по высоте (лист-презентация на 1 страницу)
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $Path)) { throw "Файл не найден: $Path" }
if ([System.IO.Path]::GetExtension($Path) -ne '.xlsx') { throw "Только .xlsx (без макросов): $Path" }
if (-not $Pdf) { $Pdf = [System.IO.Path]::ChangeExtension($Path, '.pdf') }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: макросы в открываемом файле не выполняются

try {
    $wb = $excel.Workbooks.Open($Path, 0)   # UpdateLinks=0: внешние ссылки не обновлять
    $expectedPages = 0

    foreach ($ws in $wb.Worksheets) {
        # 1) собираем строки с объединёнными ячейками — их AutoFit трогать нельзя
        $merged = @{}
        foreach ($cell in $ws.UsedRange.Cells) {
            if ($cell.MergeCells) { $merged[$cell.Row] = $true }
        }

        # 2) AutoFit по остальным строкам
        $last = $ws.UsedRange.Rows.Count + $ws.UsedRange.Row - 1
        for ($r = 1; $r -le $last; $r++) {
            if (-not $merged.ContainsKey($r)) {
                $ws.Rows.Item($r).AutoFit() | Out-Null
                if ($ws.Rows.Item($r).RowHeight -lt $MinRowHeight) {
                    $ws.Rows.Item($r).RowHeight = $MinRowHeight
                }
            }
        }

        # 3) печать: широкая таблица должна влезать в страницу по ширине
        if ($Portrait) { $ws.PageSetup.Orientation = 1 } else { $ws.PageSetup.Orientation = 2 }
        $ws.PageSetup.Zoom = $false
        $ws.PageSetup.FitToPagesWide = 1
        if ($FitTall -gt 0) { $ws.PageSetup.FitToPagesTall = $FitTall } else { $ws.PageSetup.FitToPagesTall = $false }
        $ws.PageSetup.LeftMargin   = $excel.InchesToPoints(0.3)
        $ws.PageSetup.RightMargin  = $excel.InchesToPoints(0.3)
        $ws.PageSetup.TopMargin    = $excel.InchesToPoints(0.4)
        $ws.PageSetup.BottomMargin = $excel.InchesToPoints(0.4)
        if ($TitleRow -ne '') {
            try { $ws.PageSetup.PrintTitleRows = $TitleRow } catch { }
        }

        # 4) Скрытые листы НЕ ТРОГАЕМ — заказчик прячет их намеренно.
        # Просто сообщаем, чтобы недостача страниц в PDF не выглядела потерей данных.
        if ($ws.Visible -ne -1) {
            Write-Output "$($ws.Name): СКРЫТ заказчиком — пропускаю, в PDF не войдёт"
            continue
        }

        $pages = $ws.PageSetup.Pages.Count
        if ($pages -lt 1) {
            Write-Warning "ЛИСТ '$($ws.Name)': видимый, но страниц для печати 0 — проверь PrintArea"
        }
        Write-Output "$($ws.Name): строк $last, объединённых $($merged.Count), страниц $pages"
        $expectedPages += $pages
    }

    $wb.Worksheets.Item(1).Activate()
    $wb.Save()
    $wb.ExportAsFixedFormat(0, $Pdf) | Out-Null
    $wb.Close($false)

    Write-Output "PDF: $Pdf (ожидается страниц: $expectedPages)"
    Write-Output "ПРОВЕРЬ: столько же страниц должно быть в PDF. Меньше — лист потерялся."
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
