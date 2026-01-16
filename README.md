# try-xlsx2pdf-pwsh
[Microsoft Graph]: https://learn.microsoft.com/ja-jp/graph/overview?context=graph%2Fapi%2F1.0&view=graph-rest-1.0

[Microsoft Graph] PowerShell SDK を使用して xlsx→pdf 変換のおためし

大まかな流れは次のとおりです。

1. ログイン
2. ファイルのアップロード (例： root:/FolderA/FileB.xlsx)
3. アップロードしたファイルの ID 取得
4. PDF 形式でファイルをダウンロード
5. アップロードしたファイルの削除

### Microsoft Graph PowerShell SDK のインストール

公式ドキュメント: [Install the Microsoft Graph PowerShell SDK \| Microsoft Learn](https://learn.microsoft.com/ja-jp/powershell/microsoftgraph/installation?view=graph-powershell-1.0)

```ps1
Install-Module Microsoft.Graph -Scope CurrentUser -Repository PSGallery -Force
Get-InstalledModule Microsoft.Graph
Update-Module Microsoft.Graph

```

### xlsx → pdf への変換手順

- 対話型認証で Microsoft Graph に接続します。
  - 個人用アカウントを使用する場合、TenantId は consumers または common を指定します。
- OneDrive 上の指定フォルダーに Excel ファイルをアップロードします。
  - 小さなファイルは通常の PUT でアップロードします。
  - 大きなファイルはアップロード セッションを使用してチャンク アップロードします。
- アップロードしたファイルのアイテム ID を取得します。
- format=pdf を指定してコンテンツをダウンロードし、PDF として保存します。
- 後処理として、アップロードしたファイルを削除します。
- 最後に Microsoft Graph から切断します。

> [!Important]
> - **※個人用アカウントはクライアント資格情報フロー (client_credentials フロー) に対応していません！**
> - そのため、**対話型認証（ユーザーログイン）が必須** となります。

