# =============================================================================
# RFQ 自動報價系統 v1.0
# 從 RFQ 彙整表自動計算各級距成本並產生建議售價
# =============================================================================

param(
    [string]$RfqPath = "",
    [string]$CostTablePath = "C:\Users\114D06\Documents\估價成本計算_2025.05.07更新.xlsx",
    [decimal]$ExchangeRate = 30,
    [string]$SetupMode = "production",   # production=量產, first=首次製作
    [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# =============================================================================
# 機台費率載入
# =============================================================================
function Load-MachineRates {
    param([string]$Path)

    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($Path)
    $ws = $wb.Worksheets.Item(1)

    $rates = @{}
    for ($r = 3; $r -le 40; $r++) {
        $machine   = $ws.Cells($r, 2).Text.Trim()
        $perMinRaw = $ws.Cells($r, 3).Value2
        $setupDayRaw = $ws.Cells($r, 4).Value2
        $setupHrRaw  = $ws.Cells($r, 5).Value2

        if ($machine -ne "" -and $perMinRaw -ne $null -and $perMinRaw -is [double]) {
            $rates[$machine] = @{
                PerMin      = [decimal]$perMinRaw
                SetupDayUSD = if ($setupDayRaw -is [double]) { [decimal]$setupDayRaw } else { 0 }
                SetupHrUSD  = if ($setupHrRaw  -is [double]) { [decimal]$setupHrRaw  } else { 0 }
            }
        }
    }

    $wb.Close($false); $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()

    return $rates
}

# =============================================================================
# 機台名稱 → 費率對應
# =============================================================================
function Get-MachineRate {
    param([string]$MachineText, [hashtable]$Rates)

    if (-not $MachineText) { return $null }
    $text = $MachineText.Trim()

    if ($Rates.ContainsKey($text)) { return $Rates[$text] }

    # CNC 編號（可能多台用「.」分隔，如 "CNC13.16" 或 "CNC19.22.23"）
    $cncMatches = [regex]::Matches($text, 'CNC\s*(\d+)')
    if ($cncMatches.Count -gt 0) {
        $found = @()
        foreach ($m in $cncMatches) {
            $key = "CNC$($m.Groups[1].Value)"
            if ($Rates.ContainsKey($key)) { $found += $Rates[$key] }
        }
        if ($found.Count -gt 0) {
            $avgMin = ($found | ForEach-Object { $_.PerMin } | Measure-Object -Average).Average
            $avgHr  = ($found | ForEach-Object { $_.SetupHrUSD } | Measure-Object -Average).Average
            $avgDay = ($found | ForEach-Object { $_.SetupDayUSD } | Measure-Object -Average).Average
            return @{ PerMin = [decimal]$avgMin; SetupHrUSD = [decimal]$avgHr; SetupDayUSD = [decimal]$avgDay }
        }
    }

    # 關鍵字對應
    switch -Regex ($text) {
        '走心'                  { return $Rates["CNC5"] }
        '走刀'                  { return $Rates["CNC13"] }
        'MC.*白鐵|白鐵.*MC'      { return $Rates["MC(加工白鐵)"] }
        'MC.*鋁|鋁.*MC'         { return $Rates["MC(加工鋁)"] }
        'MC.*塑膠|塑膠.*MC'      { return $Rates["MC(加工塑膠件+硬度高)"] }
        '^MC'                  { return $Rates["MC(加工白鐵)"] }
        '臥式'                  { return $Rates["臥式"] }
        '研磨|LAPPING|拋光|無心'  { return $Rates["研磨 (圓筒、拋光、無心)"] }
        '大平磨'                { return $Rates["研磨 (大平磨)"] }
        '整直'                  { if ($Rates.ContainsKey("廠內整直 ")) { return $Rates["廠內整直 "] } elseif ($Rates.ContainsKey("廠內整直")) { return $Rates["廠內整直"] } }
        '後製程|塗漆|噴塗'        { return $null }  # 後製程通常是外發或材料費，不計廠內工時
        '雷刻|雷射|雷雕'         { return $Rates["雷刻"] }
        '桌上車'                { return $Rates["桌上車床"] }
        '去毛邊|毛邊'           { return $Rates["去毛邊"] }
        '組裝|組立'             { return $Rates["組裝(無塵室組立)"] }
        '滾牙|攻牙'             { return $Rates["滾牙"] }
        '噴漆|印刷|移印'         { return $Rates["噴漆 (印刷)/移印"] }
    }

    return $null
}

# =============================================================================
# 解析 NTD 數值（彈性處理各種寫法）
# =============================================================================
function Parse-NTD {
    param([string]$Text)
    if (-not $Text) { return [decimal]0 }
    $t = $Text -replace ',', '' -replace '\s+', ' '

    if ($t -match '(\d+(?:\.\d+)?)\s*NTD\s*/\s*1?\s*(?:PC|PCS|pc|pcs|件)') { return [decimal]$Matches[1] }
    if ($t -match '(\d+(?:\.\d+)?)\s*NTD')   { return [decimal]$Matches[1] }
    if ($t -match 'NT\$\s*(\d+(?:\.\d+)?)')  { return [decimal]$Matches[1] }
    if ($t -match '(\d+(?:\.\d+)?)\s*元')    { return [decimal]$Matches[1] }
    if ($t -match '^\s*\$?\s*(\d+(?:\.\d+)?)\s*$') { return [decimal]$Matches[1] }
    return [decimal]0
}

# =============================================================================
# 解析外發廠商各級距報價
# 格式範例：「5pcs----\n150pcs----1941NTD/pcs\n400pcs----1616NTD/pcs」
# =============================================================================
function Parse-ExternalTiers {
    param([string]$Text)

    $tiers = [ordered]@{}
    if (-not $Text) { return $tiers }

    $lines = $Text -split "`r?`n|；|;"
    foreach ($line in $lines) {
        if ($line -match '(\d[\d,]*)\s*(?:pcs|PCS|pc|PC|件)\s*[-=～~＝]*\s*(\d[\d,]*(?:\.\d+)?)\s*NTD') {
            $qty = [int]($Matches[1] -replace ',', '')
            $price = [decimal]($Matches[2] -replace ',', '')
            $tiers[[string]$qty] = $price
        }
        elseif ($line -match '(\d[\d,]*)\s*(?:pcs|PCS|pc|PC|件)\s*[=＝]\s*(\d[\d,]*(?:\.\d+)?)\s*元') {
            $qty = [int]($Matches[1] -replace ',', '')
            $totalPrice = [decimal]($Matches[2] -replace ',', '')
            $tiers[[string]$qty] = [Math]::Round($totalPrice / $qty, 2)
        }
    }
    return $tiers
}

# =============================================================================
# 依數量取對應外發報價（取 <= qty 的最大級距）
# =============================================================================
function Get-TierPrice {
    param($Tiers, [int]$Qty)
    if ($Tiers.Count -eq 0) { return [decimal]0 }

    $sortedKeys = $Tiers.Keys | Sort-Object { [int]$_ }
    $best = [decimal]$Tiers[$sortedKeys[0]]
    foreach ($k in $sortedKeys) {
        if ([int]$k -le $Qty) { $best = [decimal]$Tiers[$k] }
    }
    return $best
}

# =============================================================================
# 解析設定時間（小時數）
# 範例：「打樣 4小時\n量產 4小時」「首次製作 : 8小時\n量產 : 6.5小時」
# =============================================================================
function Parse-SetupHours {
    param([string]$Text, [string]$Mode = "production")
    if (-not $Text) { return [decimal]0 }

    if ($Mode -eq "production") {
        if ($Text -match '量產[\s:：]*(\d+(?:\.\d+)?)\s*小時') { return [decimal]$Matches[1] }
        if ($Text -match '量產[\s:：]*(\d+(?:\.\d+)?)\s*h')   { return [decimal]$Matches[1] }
        if ($Text -match '量產[\s:：]*(\d+(?:\.\d+)?)\s*天')   { return [decimal]$Matches[1] * 8 }
    } else {
        if ($Text -match '(?:首次製作|首次|打樣)[\s:：]*(\d+(?:\.\d+)?)\s*小時') { return [decimal]$Matches[1] }
        if ($Text -match '(?:首次製作|首次|打樣)[\s:：]*(\d+(?:\.\d+)?)\s*天')   { return [decimal]$Matches[1] * 8 }
    }

    if ($Text -match '(\d+(?:\.\d+)?)\s*小時') { return [decimal]$Matches[1] }
    if ($Text -match '(\d+(?:\.\d+)?)\s*天')   { return [decimal]$Matches[1] * 8 }
    return [decimal]0
}

# =============================================================================
# 解析難度等級
# =============================================================================
function Parse-Difficulty {
    param([string]$Text)
    if (-not $Text) { return 0 }
    if ($Text -match '難度\s*[:：]?\s*(\d)') { return [int]$Matches[1] }
    return 0
}

# =============================================================================
# 難度 → 毛利率（取區間中間值）
# =============================================================================
function Get-MarginByDifficulty {
    param([int]$Difficulty)
    switch ($Difficulty) {
        1 { return 0.250 }
        2 { return 0.275 }
        3 { return 0.325 }
        4 { return 0.375 }
        5 { return 0.425 }
        default { return 0.300 }
    }
}

# =============================================================================
# 解析時間（分鐘）— 處理 "5.5"、"5.5分"、"15分"
# =============================================================================
function Parse-Minutes {
    param($Value)
    if ($Value -eq $null -or $Value -eq "") { return [decimal]0 }
    if ($Value -is [double] -or $Value -is [decimal] -or $Value -is [int]) { return [decimal]$Value }
    $t = [string]$Value
    if ($t -match '(\d+(?:\.\d+)?)\s*分') { return [decimal]$Matches[1] }
    if ($t -match '(\d+(?:\.\d+)?)')      { return [decimal]$Matches[1] }
    return [decimal]0
}

# =============================================================================
# 處理單一料號 Sheet
# =============================================================================
function Process-PartSheet {
    param($ws, [hashtable]$MachineRates, [decimal]$ExchangeRate, [string]$SetupMode)

    $maxRow = $ws.UsedRange.Rows.Count

    # 基本資訊
    $customer = $ws.Cells(3, 2).Text
    $partNo   = $ws.Cells(4, 4).Text
    $version  = $ws.Cells(4, 6).Text
    $material = $ws.Cells(4, 8).Text
    $surface  = $ws.Cells(4, 10).Text

    # 數量級距
    $qtys = @()
    for ($c = 4; $c -le 14; $c++) {
        $v = $ws.Cells(3, $c).Value2
        if ($v -is [double] -and $v -gt 0) { $qtys += [int]$v }
    }
    if ($qtys.Count -eq 0) { $qtys = @(1) }

    # 材料成本
    $materialCost = [decimal]0
    $materialDetails = @()
    foreach ($r in @(10, 13, 16, 19, 22)) {
        $vendorText = $ws.Cells($r, 2).Text
        $priceText  = $ws.Cells($r, 4).Text
        $price = Parse-NTD $priceText
        if ($price -gt 0) {
            $materialCost += $price
            $materialDetails += "$($vendorText -replace '\s+',' '): $price"
        }
    }

    # 外購物料
    $purchasedCost = [decimal]0
    for ($r = 25; $r -le 34; $r++) {
        $itemText  = $ws.Cells($r, 2).Text
        $priceText = $ws.Cells($r, 4).Text
        if ($itemText -ne "" -or $priceText -ne "") {
            $price = Parse-NTD $priceText
            if ($price -gt 0) { $purchasedCost += $price }
        }
    }

    # OP 行掃描
    $inHouseCostPerPc = [decimal]0
    $setupCostUSD = [decimal]0
    $maxDifficulty = 0
    $hasInHouseWork = $false
    $unmappedMachines = New-Object System.Collections.Generic.HashSet[string]
    $externalTieredCosts = [ordered]@{}
    $externalFlatCost = [decimal]0
    $opSummary = @()

    for ($r = 35; $r -le $maxRow; $r++) {
        $colA = $ws.Cells($r, 1).Text
        if (-not ($colA -match '^OP\d+$')) { continue }

        $machineText = $ws.Cells($r, 2).Text.Trim()
        $timeRaw     = $ws.Cells($r, 3).Value2
        $setupText   = $ws.Cells($r, 4).Text
        $notesText   = $ws.Cells($r, 5).Text
        $extVendor   = $ws.Cells($r, 8).Text
        $extPrice    = $ws.Cells($r, 9).Text

        $timeMin = Parse-Minutes $timeRaw

        # 廠內加工
        if ($timeMin -gt 0) {
            $hasInHouseWork = $true
            $rate = Get-MachineRate $machineText $MachineRates
            if ($rate) {
                $opCost = $timeMin * $rate.PerMin
                $inHouseCostPerPc += $opCost
                $opSummary += "$colA $machineText $($timeMin)分=NT$$([Math]::Round($opCost,1))"

                $setupHrs = Parse-SetupHours $setupText $SetupMode
                if ($setupHrs -gt 0 -and $rate.SetupHrUSD -gt 0) {
                    $setupCostUSD += $setupHrs * $rate.SetupHrUSD
                }
            } else {
                if ($machineText -ne "" -and $machineText -notmatch '^(QC|清洗|包裝|超音波|後製程|塗漆)') {
                    [void]$unmappedMachines.Add($machineText)
                }
            }
        }

        # 難度
        $diff = Parse-Difficulty $notesText
        if ($diff -gt $maxDifficulty) { $maxDifficulty = $diff }

        # 外發報價
        if ($extPrice -ne "" -or $extVendor -ne "") {
            $tiers = Parse-ExternalTiers $extPrice
            if ($tiers.Count -gt 0) {
                foreach ($k in $tiers.Keys) {
                    if (-not $externalTieredCosts.Contains($k)) { $externalTieredCosts[$k] = [decimal]0 }
                    $externalTieredCosts[$k] += $tiers[$k]
                }
            } else {
                $singlePrice = Parse-NTD $extPrice
                if ($singlePrice -gt 0) {
                    # 外發單一報價，後續會用 /0.7 計算到售價
                    $externalFlatCost += $singlePrice
                }
            }
        }
    }

    # 一次性費用（其他資訊段落）
    $oneTimeCosts = [ordered]@{}
    for ($r = 90; $r -le $maxRow; $r++) {
        $label = $ws.Cells($r, 1).Text.Trim()
        $val   = $ws.Cells($r, 2).Text
        if ($label -match '模具|治具|刀具|PIN|包材|工裝|遮噴' -and $val -ne "") {
            if ($val -match '(\d[\d,]*)\s*NTD') {
                $oneTimeCosts[$label] = [decimal]($Matches[1] -replace ',', '')
            } elseif ($val -match '(\d[\d,]*)\s*元') {
                $oneTimeCosts[$label] = [decimal]($Matches[1] -replace ',', '')
            }
        }
    }
    # 也掃描刀具段（FTA01格式：「60 PCS    10000元」）
    for ($r = 100; $r -le $maxRow; $r++) {
        $label = $ws.Cells($r, 1).Text.Trim()
        $val   = $ws.Cells($r, 2).Text
        if ($label -eq "刀具" -and $val -match '(\d+)\s*PCS?\s+(\d[\d,]*)\s*元') {
            $oneTimeCosts["刀具"] = [decimal]($Matches[2] -replace ',', '')
        }
    }

    # 決定毛利率
    $isFullyExternal = (-not $hasInHouseWork) -and ($externalTieredCosts.Count -gt 0 -or $externalFlatCost -gt 0)
    $margin = if ($isFullyExternal) {
        0.30   # 外發專案：/0.7 等同 30% 毛利
    } elseif ($maxDifficulty -gt 0) {
        Get-MarginByDifficulty $maxDifficulty
    } else {
        0.30   # 預設
    }

    $setupCostNTD = $setupCostUSD * $ExchangeRate
    $oneTimeTotal = 0
    foreach ($v in $oneTimeCosts.Values) { $oneTimeTotal += $v }

    # 計算各級距
    $qtyCosts = [ordered]@{}
    foreach ($qty in $qtys) {
        $extCost = if ($externalTieredCosts.Count -gt 0) { Get-TierPrice $externalTieredCosts $qty } else { [decimal]0 }
        $extCost += $externalFlatCost

        $setupPerPc   = if ($qty -gt 0) { [Math]::Round($setupCostNTD / $qty, 1) } else { 0 }
        $oneTimePerPc = if ($qty -gt 0 -and $oneTimeTotal -gt 0) { [Math]::Round($oneTimeTotal / $qty, 1) } else { 0 }

        # 廠內成本 → 加上設定費攤提
        $inHouseTotal = $inHouseCostPerPc + $setupPerPc

        # 全部成本
        $totalCost = $materialCost + $purchasedCost + $extCost + $inHouseTotal + $oneTimePerPc

        # 建議售價：
        # - 全外發：總成本 / 0.7
        # - 廠內加工：總成本 / (1 - 毛利率)
        $suggestedPrice = if ($isFullyExternal) {
            [Math]::Ceiling($totalCost / 0.7)
        } else {
            [Math]::Ceiling($totalCost / (1 - $margin))
        }

        $qtyCosts[[string]$qty] = @{
            Material       = [Math]::Round($materialCost, 1)
            Purchased      = [Math]::Round($purchasedCost, 1)
            External       = [Math]::Round($extCost, 1)
            InHouse        = [Math]::Round($inHouseCostPerPc, 1)
            Setup          = $setupPerPc
            OneTime        = $oneTimePerPc
            TotalCost      = [Math]::Round($totalCost, 1)
            Margin         = $margin
            SuggestedPrice = $suggestedPrice
        }
    }

    return @{
        PartNo            = $partNo
        Customer          = $customer
        Version           = $version
        Material          = $material
        Surface           = $surface
        QtyTiers          = $qtys
        Costs             = $qtyCosts
        Difficulty        = $maxDifficulty
        IsFullyExternal   = $isFullyExternal
        UnmappedMachines  = @($unmappedMachines)
        OneTimeCosts      = $oneTimeCosts
        OpSummary         = $opSummary
        MaterialDetails   = $materialDetails
        SetupCostUSD      = [Math]::Round($setupCostUSD, 2)
    }
}

# =============================================================================
# 寫入「報價彙整」工作表
# =============================================================================
function Write-SummarySheet {
    param($wb, $allParts, [decimal]$ExchangeRate, [string]$SetupMode)

    $quoteName = "報價彙整"
    # 移除既有
    for ($i = $wb.Worksheets.Count; $i -ge 1; $i--) {
        if ($wb.Worksheets.Item($i).Name -eq $quoteName) {
            $wb.Worksheets.Item($i).Delete()
        }
    }

    $ws = $wb.Worksheets.Add($wb.Worksheets.Item(1))
    $ws.Name = $quoteName

    # 標題
    $ws.Cells(1, 1) = "RFQ 報價彙整表"
    $ws.Cells(1, 1).Font.Bold = $true
    $ws.Cells(1, 1).Font.Size = 16

    $modeText = if ($SetupMode -eq "production") { "量產" } else { "首次製作" }
    $ws.Cells(2, 1) = "產生時間: $(Get-Date -Format 'yyyy/MM/dd HH:mm')  |  匯率: $ExchangeRate NTD/USD  |  設定費模式: $modeText"
    $ws.Cells(2, 1).Font.Italic = $true
    $ws.Cells(2, 1).Font.Color = 0x666666

    # 表頭
    $row = 4
    $headers = @("料號", "客戶", "版次", "材料", "難度", "毛利率", "數量",
                 "材料/件", "外購/件", "外發/件", "廠內加工/件", "設定費攤", "一次性攤",
                 "總成本/件", "建議售價/件", "級距總金額")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $ws.Cells($row, $i + 1) = $headers[$i]
        $ws.Cells($row, $i + 1).Font.Bold = $true
        $ws.Cells($row, $i + 1).Interior.Color = 0xCCE5FF
        $ws.Cells($row, $i + 1).Borders.LineStyle = 1
        $ws.Cells($row, $i + 1).HorizontalAlignment = -4108  # xlCenter
    }
    $row++

    foreach ($part in $allParts) {
        $firstRow = $row
        $tierCount = 0

        foreach ($qty in ($part.QtyTiers | Sort-Object)) {
            if (-not $part.Costs.Contains([string]$qty)) { continue }
            $c = $part.Costs[[string]$qty]
            $tierCount++

            $ws.Cells($row, 1)  = $part.PartNo
            $ws.Cells($row, 2)  = $part.Customer
            $ws.Cells($row, 3)  = $part.Version
            $ws.Cells($row, 4)  = $part.Material
            $ws.Cells($row, 5)  = if ($part.Difficulty -eq 0) { "—" } else { $part.Difficulty }
            $ws.Cells($row, 6)  = "{0:P1}" -f $c.Margin
            $ws.Cells($row, 7)  = $qty
            $ws.Cells($row, 8)  = $c.Material
            $ws.Cells($row, 9)  = $c.Purchased
            $ws.Cells($row, 10) = $c.External
            $ws.Cells($row, 11) = $c.InHouse
            $ws.Cells($row, 12) = $c.Setup
            $ws.Cells($row, 13) = $c.OneTime
            $ws.Cells($row, 14) = $c.TotalCost
            $ws.Cells($row, 15) = $c.SuggestedPrice
            $ws.Cells($row, 16) = $c.SuggestedPrice * $qty

            # 建議售價標亮
            $ws.Cells($row, 15).Font.Bold = $true
            $ws.Cells($row, 15).Font.Color = 0x0000CC
            $ws.Cells($row, 15).Interior.Color = 0xFFF2CC

            # 邊框
            for ($cc = 1; $cc -le 16; $cc++) {
                $ws.Cells($row, $cc).Borders.LineStyle = 1
            }

            $row++
        }

        # 合併同一料號的前 4 欄
        if ($tierCount -gt 1) {
            for ($mc = 1; $mc -le 5; $mc++) {
                $ws.Range($ws.Cells($firstRow, $mc), $ws.Cells($row - 1, $mc)).Merge() | Out-Null
                $ws.Range($ws.Cells($firstRow, $mc), $ws.Cells($row - 1, $mc)).VerticalAlignment = -4108
            }
        }

        # 一次性費用
        if ($part.OneTimeCosts.Count -gt 0) {
            $items = @()
            foreach ($k in $part.OneTimeCosts.Keys) {
                $v = [decimal]$part.OneTimeCosts[$k]
                $items += $k + '=NT$' + ('{0:N0}' -f $v)
            }
            $otText = "一次性費用：" + ($items -join "  /  ")
            $ws.Cells($row, 1) = $otText
            $ws.Range($ws.Cells($row, 1), $ws.Cells($row, 16)).Merge() | Out-Null
            $ws.Cells($row, 1).Font.Italic = $true
            $ws.Cells($row, 1).Font.Color = 0x666666
            $ws.Cells($row, 1).Interior.Color = 0xF5F5F5
            $row++
        }

        # 未對應機台警告
        if ($part.UnmappedMachines.Count -gt 0) {
            $ws.Cells($row, 1) = "⚠ 未對應機台（廠內加工成本可能漏算）: " + ($part.UnmappedMachines -join ", ")
            $ws.Range($ws.Cells($row, 1), $ws.Cells($row, 16)).Merge() | Out-Null
            $ws.Cells($row, 1).Font.Color = 0xCC0000
            $ws.Cells($row, 1).Interior.Color = 0xFFE5E5
            $row++
        }

        $row++  # 空行分隔
    }

    # 欄寬
    $ws.Columns("A:D").ColumnWidth = 18
    $ws.Columns("E:F").ColumnWidth = 8
    $ws.Columns("G:N").ColumnWidth = 11
    $ws.Columns("O:P").ColumnWidth = 13
    $ws.Rows("4:4").RowHeight = 30

    # 凍結窗格
    $ws.Activate()
    $ws.Application.ActiveWindow.SplitRow = 4
    $ws.Application.ActiveWindow.FreezePanes = $true
}

# =============================================================================
# 主流程
# =============================================================================
function Main {
    Write-Host ""
    Write-Host "=== RFQ 自動報價系統 ===" -ForegroundColor Cyan
    Write-Host ""

    # 自動尋找 RFQ 檔案
    if ($RfqPath -eq "") {
        $cur = Get-Location
        $candidates = Get-ChildItem -Path $cur -Filter "RFQ*.xlsx" -File -ErrorAction SilentlyContinue
        if ($candidates.Count -eq 1) {
            $script:RfqPath = $candidates[0].FullName
            Write-Host "自動偵測到 RFQ: $RfqPath" -ForegroundColor Yellow
        } elseif ($candidates.Count -gt 1) {
            Write-Host "找到多個 RFQ 檔案，請選擇：" -ForegroundColor Yellow
            for ($i = 0; $i -lt $candidates.Count; $i++) {
                Write-Host "  [$($i+1)] $($candidates[$i].Name)"
            }
            $sel = Read-Host "輸入編號"
            $script:RfqPath = $candidates[[int]$sel - 1].FullName
        } else {
            Write-Host "找不到 RFQ 檔案，請用 -RfqPath 參數指定路徑" -ForegroundColor Red
            return
        }
    }

    if (-not (Test-Path $RfqPath)) {
        Write-Host "錯誤: 找不到檔案 $RfqPath" -ForegroundColor Red
        return
    }
    if (-not (Test-Path $CostTablePath)) {
        Write-Host "錯誤: 找不到費率表 $CostTablePath" -ForegroundColor Red
        return
    }

    Write-Host "[1/3] 載入機台費率表..." -ForegroundColor Yellow
    $rates = Load-MachineRates $CostTablePath
    Write-Host "      已載入 $($rates.Count) 種機台" -ForegroundColor Green

    Write-Host "[2/3] 開啟 RFQ 彙整表..." -ForegroundColor Yellow
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($RfqPath)

    $allParts = @()
    $skipSheets = @("範本", "報價彙整")

    foreach ($ws in $wb.Worksheets) {
        if ($ws.Name -in $skipSheets) { continue }
        Write-Host "      處理料號: $($ws.Name)" -ForegroundColor White
        $part = Process-PartSheet $ws $rates $ExchangeRate $SetupMode
        $allParts += $part
    }

    Write-Host "[3/3] 產生報價彙整工作表..." -ForegroundColor Yellow
    Write-SummarySheet $wb $allParts $ExchangeRate $SetupMode

    # 儲存
    if ($OutputPath -eq "") {
        $wb.Save()
        $savedPath = $RfqPath
    } else {
        $wb.SaveAs($OutputPath)
        $savedPath = $OutputPath
    }

    $xl.Visible = $true
    $xl.ScreenUpdating = $true

    Write-Host ""
    Write-Host "✅ 報價彙整完成！" -ForegroundColor Green
    Write-Host "   檔案位置: $savedPath" -ForegroundColor Cyan
    Write-Host "   已加入「報價彙整」工作表" -ForegroundColor Cyan

    foreach ($part in $allParts) {
        if ($part.UnmappedMachines.Count -gt 0) {
            Write-Host ""
            Write-Host "⚠ 料號 $($part.PartNo) 有未對應機台：$($part.UnmappedMachines -join ', ')" -ForegroundColor Yellow
            Write-Host "  請至「報價彙整」工作表確認" -ForegroundColor Yellow
        }
    }
}

Main
