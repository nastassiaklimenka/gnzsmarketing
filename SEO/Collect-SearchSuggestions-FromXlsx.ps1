param(
    [Parameter(Mandatory = $true)]
    [string]$SeedFile,

    [string]$SeedSheet = "",
    [string]$SeedColumn = "A",
    [string]$SeedHeader = "",
    [switch]$HasHeader = $true,

    [int]$TopPerSeed = 50,

    [string]$RawOut = ".\\search_raw.csv",
    [string]$CleanOut = ".\\search_clean_clusters.csv",
    [string]$RawPhrasesOut = ".\\search_raw_phrases.txt",
    [string]$CleanPhrasesOut = ".\\search_clean_top_phrases.txt"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $SeedFile)) {
    throw "Файл не найден: $SeedFile"
}

function Convert-ExcelColumnToIndex {
    param([string]$Column)
    if ($Column -match '^\d+$') { return [int]$Column }
    $index = 0
    foreach ($ch in $Column.ToUpper().ToCharArray()) {
        $index = $index * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $index
}

function Normalize-Phrase {
    param([string]$Phrase)
    if (-not $Phrase) { return "" }
    $p = ($Phrase.ToLower() -replace '[\"`r`n`t\'']', ' ').Trim()
    $p = $p -replace '\\s+', ' '
    return $p
}

function Is-CleanPhrase {
    param([string]$Phrase)
    if ([string]::IsNullOrWhiteSpace($Phrase)) { return $false }
    if ($Phrase.Length -lt 3) { return $false }
    if ($Phrase -match '^\d{2,}$') { return $false }
    if ($Phrase -match 'https?://') { return $false }
    return $true
}

function Get-SeedPhrasesFromXlsx {
    param(
        [string]$Path,
        [string]$Sheet,
        [string]$Column,
        [string]$Header,
        [bool]$HasHeader
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $wb = $excel.Workbooks.Open((Resolve-Path $Path).Path)
        if ($Sheet) {
            $ws = $wb.Worksheets.Item($Sheet)
            if (-not $ws) { throw "Лист '$Sheet' не найден в файле $Path" }
        }
        else { $ws = $wb.Worksheets.Item(1) }

        $used = $ws.UsedRange
        $rows = $used.Rows.Count
        $cols = $used.Columns.Count

        $colIndex = 0
        if ($Header) {
            $headerRow = 1
            for ($c = 1; c -le $cols; c++) {
                $cellVal = ($ws.Cells.Item($headerRow, $c).Text).ToString().Trim()
                if ($cellVal -eq $Header) { $colIndex = $c; break }
            }
            if (-not $colIndex) { throw "Не найдено заголовок '$Header' на листе '$($ws.Name)'" }
        }
        else {
            $colIndex = Convert-ExcelColumnToIndex $Column
        }

        $startRow = 2
        if (-not $HasHeader) { $startRow = 1 }

        $items = New-Object System.Collections.Generic.HashSet[string]
        for ($r = $startRow; $r -le $rows; $r++) {
            $v = $ws.Cells.Item($r, $colIndex).Text
            if ($null -eq $v) { continue }
            $value = $v.ToString().Trim()
            if (-not $value) { continue }
            [void]$items.Add($value)
        }

        if ($items.Count -eq 0) {
            throw "Не найдено seed-фраз в $Path (лист: $($ws.Name), колонка: $Column)"
        }
        return @($items)
    }
    finally {
        if ($wb) { $wb.Close($false) }
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Get-YandexSuggestions {
    param([string]$Phrase)

    $uri = 'https://suggest.yandex.ru/suggest-ff.cgi?part=' + [uri]::EscapeDataString($Phrase) + '&v=4&partn=2&lang=ru&callback=suggest'
    $resp = Invoke-WebRequest -Uri $uri -UseBasicParsing -ErrorAction Stop
    $text = $resp.Content
    if ($text -notmatch '^suggest\((.*)\)$') {
        return @()
    }

    $json = $Matches[1]
    $obj = $json | ConvertFrom-Json

    # структура: ["seed", ["фраза1","фраза2",...], ...]
    if ($obj.Count -lt 2 -or -not ($obj[1] -is [System.Array])) { return @() }
    return [string[]]$obj[1]
}

$seedPhrases = Get-SeedPhrasesFromXlsx -Path $SeedFile -Sheet $SeedSheet -Column $SeedColumn -Header $SeedHeader -HasHeader $HasHeader
Write-Host "Загружено seed-фраз: $($seedPhrases.Count)"

$rawRows = New-Object System.Collections.Generic.List[object]

foreach ($seed in $seedPhrases) {
    Write-Host "Обрабатывается: $seed"
    try {
        $suggestions = Get-YandexSuggestions -Phrase $seed
        if ($suggestions -and $suggestions.Count -gt 0) {
            $rank = 1
            foreach ($s in $suggestions) {
                $rawRows.Add([pscustomobject]@{
                    source = 'yandex_suggest'
                    seed = $seed
                    phrase = $s
                    query = $s
                    rank = $rank
                    searchCount = $null
                })
                $rank++
            }
        }
    }
    catch {
        Write-Warning "Не удалось получить подсказки для '$seed': $($_.Exception.Message)"
    }
}

$rawRows | Export-Csv -Path $RawOut -Encoding UTF8 -NoTypeInformation

$cleanRows = New-Object System.Collections.Generic.List[object]
$grouped = $rawRows | Group-Object seed

foreach ($g in $grouped) {
    $seed = $g.Name
    $dedup = @{}

    foreach ($r in $g.Group) {
        $norm = Normalize-Phrase $r.query
        if (-not (Is-CleanPhrase $norm)) { continue }

        if (-not $dedup.ContainsKey($norm)) {
            $dedup[$norm] = [pscustomobject]@{
                seed = $seed
                query = $r.query
                phrase = $r.query
                rank = [int]$r.rank
                searchCount = $null
                source = $r.source
            }
        }
    }

    $topRows = $dedup.Values |
        Sort-Object rank |
        Select-Object -First $TopPerSeed

    foreach ($row in $topRows) {
        $cleanRows.Add([pscustomobject]@{
            seed = $row.seed
            query = $row.query
            phrase = $row.phrase
            rank = $row.rank
            searchCount = $row.searchCount
            source = $row.source
        })
    }
}

$cleanRows | Export-Csv -Path $CleanOut -Encoding UTF8 -NoTypeInformation

$rawRows |
    Select-Object -ExpandProperty query |
    Sort-Object -Unique |
    Out-File -FilePath $RawPhrasesOut -Encoding utf8

$cleanRows |
    Select-Object -ExpandProperty query |
    Sort-Object -Unique |
    Out-File -FilePath $CleanPhrasesOut -Encoding utf8

Write-Host "Готово:`n 1) Сырой список: $RawOut`n 2) Чистые кластеры: $CleanOut`n 3) Сырые фразы: $RawPhrasesOut`n 4) Топ фразы на seed: $CleanPhrasesOut"
