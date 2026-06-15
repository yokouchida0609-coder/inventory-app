# ============================================================
# テレビ番組キーワードチェッカー (NHK API版) - LINE通知スクリプト
# ============================================================

$LINE_TOKEN   = "/kKV3oJmZMO53gn5RHqB11rhKK9+oB4RVl8H4NcYiD2W+q9UayP9pugthSz1UDyLkoGCJqE4hAxZfGMR/QT752WYxsu16pjnxbsqLPH12rV58sPvahXSINbxylyYf8xD8uio7KivMicUAPxvnljh5gdB04t89/1O/w1cDnyilFU="
$LINE_USER    = "U33e8e646619bfebe28509280c3a5be2a"
$NHK_API_KEY  = "mu7suWNCXboDp8GOoNvkBW1aJHAD3CLW"
$AREA         = "130"   # 東京
$SCRIPT_DIR   = Split-Path -Parent $MyInvocation.MyCommand.Path
$KEYWORDS_FILE  = "$SCRIPT_DIR\tv_keywords.json"
$EXCLUDE_FILE   = "$SCRIPT_DIR\tv_exclude.json"
$NOTIFIED_FILE  = "$SCRIPT_DIR\tv_notified.json"
$DAYS_AHEAD   = 1   # 前日に通知

# NHKサービス一覧
$SERVICES = @(
    @{ id = "g1"; name = "NHK総合" },
    @{ id = "g2"; name = "NHK Eテレ" },
    @{ id = "s1"; name = "NHK BS1" },
    @{ id = "s4"; name = "NHK BS" }
)

# ============================================================
# LINE通知
# ============================================================
function Send-LineMessage($message) {
    $headers = @{ "Authorization" = "Bearer $LINE_TOKEN" }
    $body = @{ "to" = $LINE_USER; "messages" = @(@{ "type" = "text"; "text" = $message }) }
    try {
        Invoke-RestMethod -Uri "https://api.line.me/v2/bot/message/push" `
            -Method Post -Headers $headers `
            -ContentType "application/json; charset=utf-8" `
            -Body ([System.Text.Encoding]::UTF8.GetBytes(($body | ConvertTo-Json -Depth 5)))
    } catch {
        Write-Host "LINE通知エラー: $_"
    }
}

# ============================================================
# 通知済みリスト読み込み・保存
# ============================================================
function Get-Notified {
    if (Test-Path $NOTIFIED_FILE) {
        $data = Get-Content $NOTIFIED_FILE -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($data) { return @($data) }
    }
    return @()
}

function Save-Notified($list) {
    $list | ConvertTo-Json | Set-Content $NOTIFIED_FILE -Encoding UTF8
}

# ============================================================
# NHK APIから番組表取得
# ============================================================
function Get-NHKPrograms($date) {
    $dateStr = $date.ToString("yyyy-MM-dd")
    $programs = @()

    foreach ($svc in $SERVICES) {
        $url = "https://program-api.nhk.jp/v3/papiPgDateTv?service=$($svc.id)&area=$AREA&date=$dateStr&key=$NHK_API_KEY"
        try {
            $wr = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30
            $json = [System.Text.Encoding]::UTF8.GetString($wr.RawContentStream.ToArray())
            $response = $json | ConvertFrom-Json
            $items = $response.($svc.id).publication
            if ($items) {
                foreach ($item in $items) {
                    $programs += @{
                        Title    = $item.name
                        Subtitle = $item.description
                        Time     = $item.startDate
                        Channel  = $svc.name
                    }
                }
            }
        } catch {
            Write-Host "NHK API エラー ($($svc.name)): $_"
        }
    }
    return $programs
}

# ============================================================
# メイン処理
# ============================================================
$keywords = Get-Content $KEYWORDS_FILE -Encoding UTF8 | ConvertFrom-Json
$excludes = if (Test-Path $EXCLUDE_FILE) { Get-Content $EXCLUDE_FILE -Encoding UTF8 | ConvertFrom-Json } else { @() }
$notified = Get-Notified
$targetDate = (Get-Date).AddDays($DAYS_AHEAD)
$targetDateStr = $targetDate.ToString("yyyy/MM/dd")

Write-Host "チェック対象日: $targetDateStr"
Write-Host "キーワード: $($keywords -join ', ')"

$programs = Get-NHKPrograms $targetDate
Write-Host "取得番組数: $($programs.Count)"

$hits = @()
foreach ($prog in $programs) {
    $searchText = "$($prog.Title) $($prog.Subtitle)"

    # 除外ワードチェック
    $excluded = $false
    foreach ($ex in $excludes) {
        if ($prog.Title -match [regex]::Escape($ex)) {
            $excluded = $true
            break
        }
    }
    if ($excluded) { continue }

    foreach ($kw in $keywords) {
        if ($searchText -match [regex]::Escape($kw)) {
            $id = "$($prog.Time)_$($prog.Title)_$kw"
            if ($notified -notcontains $id) {
                $hits += @{ Program = $prog; Keyword = $kw; Id = $id }
                $notified += $id
            }
            break
        }
    }
}

if ($hits.Count -eq 0) {
    Write-Host "ヒットなし"
} else {
    $msg = "[TV番組アラート] 明日の放送`n"
    $msg += "----------------`n"
    foreach ($h in $hits) {
        $timeStr = ""
        if ($h.Program.Time) {
            $timeStr = ([datetime]$h.Program.Time).ToString("MM/dd HH:mm")
        }
        $msg += "KW:[$($h.Keyword)]`n"
        $msg += "$timeStr $($h.Program.Channel)`n"
        $msg += "$($h.Program.Title)`n"
        if ($h.Program.Subtitle) { $msg += "($($h.Program.Subtitle))`n" }
        $msg += "----------------`n"
    }
    Write-Host $msg
    Send-LineMessage $msg
    Save-Notified $notified
    Write-Host "$($hits.Count)件を通知しました！"
}
