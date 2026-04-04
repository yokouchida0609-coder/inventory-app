﻿# ============================================================
# 株価下落アラート - LINE通知スクリプト
# ============================================================

$LINE_TOKEN  = "/kKV3oJmZMO53gn5RHqB11rhKK9+oB4RVl8H4NcYiD2W+q9UayP9pugthSz1UDyLkoGCJqE4hAxZfGMR/QT752WYxsu16pjnxbsqLPH12rV58sPvahXSINbxylyYf8xD8uio7KivMicUAPxvnljh5gdB04t89/1O/w1cDnyilFU="
$LINE_USER   = "U33e8e646619bfebe28509280c3a5be2a"
$THRESHOLD   = 0.05   # 5% 下落でアラート
$SCRIPT_DIR  = Split-Path -Parent $MyInvocation.MyCommand.Path
$BASELINE_FILE = "$SCRIPT_DIR\stock_baseline.json"

# 監視銘柄リスト（コード: 銘柄名）
$STOCKS = [ordered]@{
    "9986" = "蔵王産業"
    "3076" = "あいHD"
    "8130" = "サンゲツ"
    "2659" = "サンエー"
    "3333" = "あさひ"
    "4008" = "住友精化"
    "4042" = "東ソー"
    "4097" = "高圧ガス工業"
    "8309" = "三井住友トラストグループ"
    "8725" = "MS&ADインシュアランスグループHD"
    "8593" = "三菱HCキャピタル"
    "8584" = "ジャックス"
    "6785" = "鈴木"
    "7723" = "愛知時計電機"
    "3231" = "野村不動産HD"
    "3003" = "ヒューリック"
    "2169" = "CDS"
    "9757" = "船井総研HD"
    "9769" = "学究社"
    "4641" = "アルプス技研"
    "3817" = "SRAホールディングス"
    "3901" = "マークラインズ"
    "4674" = "クレスコ"
    "2003" = "日東富士製粉"
    "1414" = "ショーボンドHD"
    "1928" = "積水ハウス"
    "6345" = "アイチコーポレーション"
    "9364" = "上組"
    "9381" = "エーアイティー"
    "5388" = "クニミネ工業"
    "7989" = "立川ブラインド工業"
    "7820" = "ニホンフラッシュ"
    "7994" = "オカムラ"
    "4540" = "ツムラ"
    "1343" = "NF・J-REIT ETF"
    "1925" = "大和ハウス工業"
}

# ---- 株価取得 ----
function Get-StockPrice($code) {
    try {
        $url = "https://query1.finance.yahoo.com/v8/finance/chart/$code.T?interval=1d&range=1d"
        $headers = @{ "User-Agent" = "Mozilla/5.0" }
        $res = Invoke-RestMethod -Uri $url -Headers $headers -TimeoutSec 10
        $price = $res.chart.result[0].meta.regularMarketPrice
        return [math]::Round($price, 0)
    } catch {
        return $null
    }
}

# ---- LINE送信 ----
function Send-Line($message) {
    $url = "https://api.line.me/v2/bot/message/push"
    $headers = @{
        "Content-Type"  = "application/json"
        "Authorization" = "Bearer $LINE_TOKEN"
    }
    $body = @{
        to       = $LINE_USER
        messages = @(@{ type = "text"; text = $message })
    } | ConvertTo-Json -Depth 5
    try {
        $res = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -TimeoutSec 10
        Write-Host "LINE送信成功"
    } catch {
        Write-Host "LINE送信失敗: $_"
    }
}

# ---- メイン処理 ----
$now = Get-Date -Format "yyyy/MM/dd HH:mm"
Write-Host "[$now] 株価チェック開始"

# 基準価格の読み込み
if (Test-Path $BASELINE_FILE) {
    $baseline = Get-Content $BASELINE_FILE -Raw -Encoding UTF8 | ConvertFrom-Json
} else {
    $baseline = [PSCustomObject]@{}
}

$alerts = @()
$updated = $false

foreach ($code in $STOCKS.Keys) {
    $name  = $STOCKS[$code]
    $price = Get-StockPrice $code

    if ($null -eq $price) {
        Write-Host "  [$code] $name : 取得失敗"
        continue
    }

    $basePrice = $baseline.$code

    if ($null -eq $basePrice) {
        # 初回: 基準価格を登録
        $baseline | Add-Member -NotePropertyName $code -NotePropertyValue $price -Force
        $updated = $true
        Write-Host "  [$code] $name : 基準価格登録 ${price}円"
    } else {
        $change = ($price - $basePrice) / $basePrice
        $changeStr = "{0:P1}" -f $change
        Write-Host "  [$code] $name : ${price}円 (基準 ${basePrice}円, $changeStr)"

        if ($change -le -$THRESHOLD) {
            $alerts += "⚠️ $name（$code）`n  基準: ${basePrice}円 → 現在: ${price}円`n  下落率: $changeStr"
        }
    }
}

# 基準価格を保存
if ($updated) {
    $baseline | ConvertTo-Json | Set-Content $BASELINE_FILE -Encoding UTF8
    Write-Host "基準価格を保存しました"
}

# アラート送信
if ($alerts.Count -gt 0) {
    $message = "📉 株価下落アラート [$now]`n`n" + ($alerts -join "`n`n")
    Write-Host "アラート送信: $($alerts.Count)件"
    Send-Line $message
} else {
    Write-Host "[$now] 下落銘柄なし"
}
