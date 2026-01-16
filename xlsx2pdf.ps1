# Microsoft Graph PowerShell による OneDrive 操作スクリプト

# 設定値
$TargetFolder = "FolderA"                                   # OneDrive上のフォルダ名
$TargetFileName = "FileB.xlsx"                              # OneDrive上のファイル名
$SourceFile = "FileB.xlsx"                  # アップロード元ファイル
$DownloadPath = "FileB.pdf"                 # PDFダウンロード先

# =============================================================================
# 1. Microsoft Graph への接続
# =============================================================================
Write-Host "1. Microsoft Graph に接続中..." -ForegroundColor Cyan

# 必要なスコープを指定
$Scopes = @(
    "Files.ReadWrite",
    "Files.ReadWrite.All",
    "User.Read"
)

# 【重要】個人用アカウントの場合は TenantId を 'consumers' または 'common' に指定
$TenantId = "consumers" 
try {
    # 対話型認証でログイン
    Connect-MgGraph -TenantId $TenantId -Scopes $Scopes
    Write-Host "✓ 接続成功" -ForegroundColor Green
} catch {
    Write-Host "✗ 接続失敗: $_" -ForegroundColor Red
    exit
}

# 現在のコンテキスト確認
$context = Get-MgContext
Write-Host "接続ユーザー: $($context.Account)" -ForegroundColor Yellow

# =============================================================================
# 2. ファイルのアップロード
# =============================================================================
Write-Host "`n2. ファイルをアップロード中..." -ForegroundColor Cyan

try {
    # ローカルファイルの存在確認
    if (-not (Test-Path $SourceFile)) {
        throw "アップロード元ファイルが見つかりません: $SourceFile"
    }

    # ファイル内容を読み込み
    $fileContent = [System.IO.File]::ReadAllBytes($SourceFile)
    $fileSize = (Get-Item $SourceFile).Length
    Write-Host "ファイルサイズ: $([math]::Round($fileSize / 1MB, 2)) MB" -ForegroundColor Yellow

    # アップロード先のパス
    $uploadPath = "https://graph.microsoft.com/v1.0/me/drive/root:/$TargetFolder/${TargetFileName}:/content"

    # ファイルをアップロード
    if ($fileSize -lt 4MB) {
        # 小さいファイルは直接アップロード
        Invoke-MgGraphRequest -Method PUT -Uri $uploadPath -Body $fileContent -ContentType "application/octet-stream"
    } else {
        # 大きいファイルはアップロードセッションを使用
        $uploadSessionUri = "v1.0/me/drive/root:/$TargetFolder/${TargetFileName}:/createUploadSession"
        $uploadSession = Invoke-MgGraphRequest -Method POST -Uri $uploadSessionUri -Body @{}
        
        # チャンクアップロード処理
        $uploadUrl = $uploadSession.uploadUrl
        $chunkSize = 320KB * 10  # 3.2MB
        $offset = 0
        
        while ($offset -lt $fileSize) {
            $chunkLength = [Math]::Min($chunkSize, $fileSize - $offset)
            $chunk = $fileContent[$offset..($offset + $chunkLength - 1)]
            
            $headers = @{
                "Content-Length" = $chunkLength
                "Content-Range" = "bytes $offset-$($offset + $chunkLength - 1)/$fileSize"
            }
            
            Invoke-RestMethod -Method PUT -Uri $uploadUrl -Body $chunk -Headers $headers
            $offset += $chunkLength
            
            $progress = [math]::Round(($offset / $fileSize) * 100, 1)
            Write-Host "アップロード進捗: $progress%" -ForegroundColor Yellow
        }
    }

    Write-Host "✓ アップロード成功" -ForegroundColor Green
} catch {
    Write-Host "✗ アップロード失敗: $_" -ForegroundColor Red
    Disconnect-MgGraph
    exit
}

# =============================================================================
# 3. アップロードファイルのID取得
# =============================================================================
Write-Host "`n3. ファイルIDを取得中..." -ForegroundColor Cyan

try {
    $itemPath = "https://graph.microsoft.com/v1.0/me/drive/root:/$TargetFolder/${TargetFileName}"
    $fileItem = Invoke-MgGraphRequest -Method GET -Uri $itemPath
    
    $fileId = $fileItem.id
    Write-Host "✓ ファイルID: $fileId" -ForegroundColor Green
} catch {
    Write-Host "✗ ファイルID取得失敗: $_" -ForegroundColor Red
    Disconnect-MgGraph
    exit
}

# =============================================================================
# 4. PDF形式でファイルをダウンロード
# =============================================================================
Write-Host "`n4. PDFとしてダウンロード中..." -ForegroundColor Cyan

try {
    # PDF変換ダウンロードのエンドポイント
    $convertUri = "https://graph.microsoft.com/v1.0/me/drive/items/$fileId/content?format=pdf"
    
    # PDFをダウンロード
    $pdfContent = Invoke-MgGraphRequest -Method GET -Uri $convertUri -OutputType HttpResponseMessage
    
    # レスポンスからバイトストリームを取得
    $stream = $pdfContent.Content.ReadAsStreamAsync().Result
    $fileStream = [System.IO.File]::Create($DownloadPath)
    $stream.CopyTo($fileStream)
    $fileStream.Close()
    $stream.Close()
    
    Write-Host "✓ PDFダウンロード成功: $DownloadPath" -ForegroundColor Green
} catch {
    Write-Host "✗ PDFダウンロード失敗: $_" -ForegroundColor Red
    Write-Host "注意: Excelファイルは変換可能ですが、他の形式ではエラーになる場合があります" -ForegroundColor Yellow
}

# =============================================================================
# 5. アップロードファイルの削除
# =============================================================================
Write-Host "`n5. ファイルを削除中..." -ForegroundColor Cyan

try {
    $deleteUri = "https://graph.microsoft.com/v1.0/me/drive/items/$fileId"
    Invoke-MgGraphRequest -Method DELETE -Uri $deleteUri
    
    Write-Host "✓ ファイル削除成功" -ForegroundColor Green
} catch {
    Write-Host "✗ ファイル削除失敗: $_" -ForegroundColor Red
}

# =============================================================================
# 切断
# =============================================================================
Write-Host "`n処理完了。Microsoft Graph から切断します。" -ForegroundColor Cyan
Disconnect-MgGraph
Write-Host "✓ 切断完了" -ForegroundColor Green
