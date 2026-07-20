# Microsoft Graph Excel MCP Server

SharePoint サイトのドキュメントライブラリ上にある Excel ファイルを、MCP クライアントから操作するためのサーバーです。

このサーバーは Microsoft Graph API を使って Excel ファイルを操作します。初回利用時に Microsoft アカウントでサインインすると、以後は保存されたトークンを使って MCP tool を実行できます。

## できること

- SharePoint サイトの既定ドライブに Excel ファイルを作成する
- ファイルをアップロード、ダウンロードする
- ワークシートを一覧取得、作成、名前変更する
- セルや範囲に値を書き込む
- セルや範囲の値を読み取る
- 行を追加、挿入、更新する
- Excel workbook session を開始、終了する

## 事前準備

### 1. Azure アプリ登録を用意する

Microsoft Entra ID のアプリ登録で、Microsoft Graph の delegated permission を追加します。

最低限、次の権限を設定してください。

```json
[
  "Sites.ReadWrite.All",
  "Files.ReadWrite.All"
]
```

組織の設定によっては管理者の同意が必要です。

### 2. Redirect URI を登録する

ローカルで `http://localhost:3001` として起動する場合、Azure アプリ登録の redirect URI に次を追加します。

```text
http://localhost:3001/auth/callback
```

Azure アプリ登録の種類により、設定方法が変わります。

- Web プラットフォームに redirect URI を登録する場合: client secret が必要です
- Mobile and desktop applications として登録する場合: client secret は空で使えます

### 3. appsettings.json を作成する

`appsettings.json.template` を参考に、`appsettings.json` を作成します。

```json
{
  "Settings": {
    "ClientId": "Azure アプリ登録のアプリケーション クライアント ID",
    "ClientSecret": "Web プラットフォームを使う場合の client secret。不要な場合は空文字",
    "TenantId": "Microsoft Entra ID のテナント ID",
    "GraphUserScopes": [
      "Sites.ReadWrite.All",
      "Files.ReadWrite.All"
    ],
    "TokenCachePath": ""
  }
}
```

`ClientSecret` には、Azure portal で作成したシークレットの **Value** を指定します。Secret ID ではありません。

`TokenCachePath` を空にした場合、トークンキャッシュは OS のローカルアプリケーションデータ配下に profile ごとに保存されます。明示した場合も、`default` 以外の profile はファイル名に profile 名を付けて分離されます。

## 起動

```bash
dotnet run --urls http://localhost:3001
```

MCP endpoint は次の URL です。

```text
http://localhost:3001/mcp
```

ブラウザーで `/` を開くと、MCP endpoint、認証 endpoint、profile 指定方法を確認できます。

```text
http://localhost:3001/
```

## 認証

初回利用時、ブラウザーで認証 URL を開いて Microsoft にサインインします。

```text
http://localhost:3001/auth/login
```

認証状態は次の URL で確認できます。

```text
http://localhost:3001/auth/status
```

認証済みの場合は、次のような JSON が返ります。

```json
{
  "profile": "default",
  "authenticated": true
}
```

## 複数ユーザーで使う場合

このサーバーは `profile` ごとに Microsoft Graph の認証情報、トークンキャッシュ、Excel workbook session を分離します。

ユーザーごとに別の profile 名で認証してください。

```text
http://localhost:3001/auth/login?profile=alice
http://localhost:3001/auth/login?profile=bob
```

MCP クライアントから tool を呼び出すときは、HTTP ヘッダー `X-Excel-Mcp-Profile` に同じ profile 名を指定します。

```text
X-Excel-Mcp-Profile: alice
```

ヘッダーがない場合は `default` profile を使います。

profile 名に使える文字は、英数字、ハイフン、アンダースコア、ドットです。

Dify などでエンドユーザーごとに profile を変えたい場合は、Dify 側のユーザー ID を `X-Excel-Mcp-Profile` ヘッダーに設定してください。例:

```text
X-Excel-Mcp-Profile: {{#sys.user_id#}}
```

Dify の MCP 設定で動的ヘッダーを設定できない場合は、Dify とこの MCP サーバーの間に proxy を置き、proxy 側でユーザー ID から `X-Excel-Mcp-Profile` を付与してください。

## SharePoint siteId

各 tool では SharePoint siteId を指定します。形式は次の通りです。

```text
[テナント名].sharepoint.com,[Site GUID],[Web GUID]
```

Site GUID と Web GUID は SharePoint の REST API で確認できます。

```text
https://[テナント名].sharepoint.com/sites/[サイト名]/_api/site/id
https://[テナント名].sharepoint.com/sites/[サイト名]/_api/web/id
```

## MCP tools

公開している MCP tool は次の通りです。

| Tool | 説明 |
| --- | --- |
| `get_site_default_drive` | SharePoint サイトの既定ドライブを取得します |
| `upload_file_base64` | base64 文字列からファイルをアップロードします |
| `download_file_base64` | ファイルを base64 文字列としてダウンロードします |
| `create_empty_excel_file` | 空の `.xlsx` ファイルを作成します |
| `get_drive_item_by_path` | パスから DriveItem を取得します |
| `create_excel_session` | Excel workbook session を開始します |
| `close_excel_session` | Excel workbook session を終了します |
| `list_worksheets` | ワークシート一覧を取得します |
| `create_worksheet` | ワークシートを作成します |
| `rename_worksheet` | ワークシート名を変更します |
| `set_cell_value` | 単一セルに値を書き込みます |
| `get_cell_value` | 単一セルの値を取得します |
| `set_range_values` | 範囲に二次元配列の値を書き込みます |
| `get_used_range` | ワークシートの使用範囲を取得します |
| `get_range_values` | 指定範囲の値を取得します |
| `add_row` | 使用範囲の末尾に 1 行追加します |
| `add_rows` | 使用範囲の末尾に複数行追加します |
| `insert_row` | 指定行に 1 行挿入します |
| `insert_rows` | 指定行に複数行挿入します |
| `update_row` | 指定行を更新します |

各 tool は Graph SDK のオブジェクトを直接受け取らず、`siteId`、`filePath`、`worksheetName`、`cellAddress`、`rangeAddress`、`values` などの JSON 化しやすい引数で呼び出します。

## よくあるエラー

### AADSTS50011: redirect URI does not match

認証リクエストで送っている redirect URI が、Azure アプリ登録に登録されていません。

ローカルで `http://localhost:3001` 起動している場合は、Azure アプリ登録に次を完全一致で追加してください。

```text
http://localhost:3001/auth/callback
```

### AADSTS7000218: client_secret が必要

Azure アプリ登録で Web プラットフォームの redirect URI を使っている場合、token 交換には client secret が必要です。

`appsettings.json` の `ClientSecret` に、Azure portal で作成した client secret の Value を設定してください。

### Microsoft Graph is not authenticated

該当 profile がまだ認証されていません。

```text
http://localhost:3001/auth/login?profile=<profile名>
```

MCP tool 呼び出しで `X-Excel-Mcp-Profile` を指定している場合は、同じ profile 名で認証してください。
