param(
    [Parameter(Mandatory = $true)]
    [string]$FolderId,

    [Parameter(Mandatory = $true)]
    [string]$SeedFile,

    [string]$SeedSheet = "",
    [string]$SeedColumn = "A",
    [string]$SeedHeader = "",
    [switch]$HasHeader = $true,

    [int]$NumPhrases = 200,
    [int]$TopPerSeed = 50,
    [string[]]$Regions = @("213"),

    [string]$RawOut = ".\\wordstat_raw.csv",
    [string]$CleanOut = ".\\wordstat_clean_clusters.csv",
    [string]$RawPhrasesOut = ".\\wordstat_raw_phrases.txt",
    [string]$CleanPhrasesOut = ".\\wordstat_clean_top_phrases.txt",

    [string]$ApiKey = $env:YANDEX_API_KEY
)

$ErrorActionPreference = "Stop"

if (-not $ApiKey) {
    throw "Не задан API-ключ. Установи: `$env:YANDEX_API_KEY = 'ключ'"
}
if (-not (Test-Path $SeedFile)) {
    throw "Файл не найден: $SeedFile"
}

function Convert-ExcelColumnToIndex {
    param([string]$Column)

    $c = $Column.Trim().ToUpper()
    if ($c -match '^\d+$') { return [int]$c }

    $index = 0
    $chars = $c.ToCharArray()
    foreach ($ch in $chars) {
        $index = $index * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $index
}

function Normalize-Phrase {
    param([string]$Phrase)
    if (-not $Phrase) { return "" }
    $p = ($Phrase.ToLower() -replace '["`r`n`t' + "'" + ']', ' ').Trim()
    $p = $p -replace '\s+', ' '
    return $p
}

function Is-CleanPhrase {
    param([string]$Phrase)
    if ([string]::IsNullOrWhiteSpace($Phrase)) { return $false }
    $badPatterns = @(
        '^\d{2,}$',
        'https?://',
        '[a-z]{2,}\\.ru',
        '^\\W+$'
    )
    foreach ($pat in $badPatterns) {
        if ($Phrase -match $pat) { return $false }
    }
    if ($Phrase.Length -lt 3) { return $false }
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
            if (-not $ws) {
                throw "Лист '$Sheet' не найден в файле $Path"
            }
        } else {
            $ws = $wb.Worksheets.Item(1)
        }

        $used = $ws.UsedRange
        $rows = $used.Rows.Count
        $cols = $used.Columns.Count

        $colIndex = 0
        if ($Header) {
            $headerRow = 1
            for ($c = 1; $c -le $cols; $c++) {
                $cellVal = ($ws.Cells.Item($headerRow, $c).Text).ToString().Trim()
                if ($cellVal -eq $Header) {
                    $colIndex = $c
                    break
                }
            }
            if (-not $colIndex) {
                throw "Не найдено заголовок '$Header' на листе '$($ws.Name)'"
            }
        } else {
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

function Get-WordstatTop {
    param(
        [string]$Phrase,
        [string[]]$Regions,
        [int]$NumPhrases,
        [string]$FolderId,
        [hashtable]$Headers,
        [string]$Uri
    )

    $body = @{
        phrase     = $Phrase
        numPhrases = $NumPhrases
        regions    = $Regions
        devices    = @('DEVICE_ALL')
        folderId   = $FolderId
    } | ConvertTo-Json -Depth 10

    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            return Invoke-RestMethod -Method Post -Uri $Uri -Headers $Headers -Body $body -ContentType 'application/json'
        }
        catch {
            $code = $_.Exception.Response.StatusCode.value__
            if ($code -eq 429) {
                Start-Sleep -Seconds ([Math]::Min(60, (5 * [Math]::Pow(2, $attempt))))
                continue
            }
            throw
        }
    }

    throw "Не удалось получить данные по '$Phrase' после 5 попыток."
}

$uri = 'https://searchapi.api.cloud.yandex.net/v2/wordstat/topRequests'
$headers = @{
    Authorization = "Api-key $ApiKey"
    'Content-Type' = 'application/json'
}

$seedPhrases = Get-SeedPhrasesFromXlsx -Path $SeedFile -Sheet $SeedSheet -Column $SeedColumn -Header $SeedHeader -HasHeader $HasHeader

Write-Host "Загружено seed-фраз: $($seedPhrases.Count)"

$rawRows = New-Object System.Collections.Generic.List[object]

foreach ($seed in $seedPhrases) {
    Write-Host "Обрабатывается: $seed"
    $resp = Get-WordstatTop -Phrase $seed -Regions $Regions -NumPhrases $NumPhrases -FolderId $FolderId -Headers $headers -Uri $uri

    foreach ($it in $resp.results) {
        $rawRows.Add([pscustomobject]@{
            source      = 'results'
            seed        = $seed
            phrase      = $it.phrase
            query       = $it.phrase
            searchCount = [int64]$it.count
        })
    }

    if ($resp.associations) {
        foreach ($it in $resp.associations) {
            $rawRows.Add([pscustomobject]@{
                source      = 'associations'
                seed        = $seed
                phrase      = $it.phrase
                query       = $it.phrase
                searchCount = [int64]$it.count
            })
        }
    }

    Start-Sleep -Milliseconds 300
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

        if ((-not $dedup.ContainsKey($norm)) -or ($r.searchCount -gt $dedup[$norm].searchCount)) {
            $dedup[$norm] = [pscustomobject]@{
                seed        = $seed
                query       = $r.query
                phrase      = $r.query
                searchCount = [int64]$r.searchCount
                source      = $r.source
            }
        }
    }

    $topRows = $dedup.Values |
        Sort-Object searchCount -Descending |
        Select-Object -First $TopPerSeed

    foreach ($row in $topRows) {
        $cleanRows.Add([pscustomobject]@{
            seed        = $row.seed
            query       = $row.query
            phrase      = $row.phrase
            searchCount = $row.searchCount
            source      = $row.source
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

Write-Host "Готово:`n 1) Сырой список:               $RawOut`n 2) Чистые кластеры:             $CleanOut`n 3) Сырые фразы (txt):            $RawPhrasesOut`n 4) Топ фразы на каждый seed (txt):  $CleanPhrasesOut"
