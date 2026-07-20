# エクセル編集

Microsoft Graph API を使って、SharePoint サイトのドキュメントライブラリ上にある Excel ファイルを操作するサンプルです。

## できること

- SharePoint サイトのデフォルトドライブを取得する
- ドライブにファイルをアップロードする
- 埋め込みテンプレートから空の Excel ファイルを作成する
- Excel 編集セッションを開始、終了する
- ワークシート一覧を取得する
- ワークシートを作成する
- ワークシート名を変更する
- 行を追加、挿入、更新する
- 指定セルに値を設定、取得する
- 始点セルと二重配列から Range に値を設定する
- 指定ワークシートの有効な Range を取得する
- 指定 Range 内の全セル値を取得する

## 必要な権限

`appsettings.json` の `GraphUserScopes` に、少なくとも以下の Microsoft Graph delegated permission を設定します。

```json
[
  "Sites.ReadWrite.All",
  "Files.ReadWrite.All"
]
```

Azure portal 側でも同じ delegated permission を追加し、必要に応じて管理者の同意を付与してください。

## 設定

`appsettings.json` を作成し、Azure AD アプリ登録の値を設定します。Azure portal のアプリ登録では、リダイレクト URI に MCP サーバーの callback URL を登録してください。ローカルで `http://localhost:3001` 起動する場合は `http://localhost:3001/auth/callback` です。

Web プラットフォームに redirect URI を登録した場合は client secret が必要です。Mobile and desktop applications の public client として登録した場合は client secret を空にできます。

```json
{
  "Settings": {
    "ClientId": "アプリケーション クライアント ID",
    "ClientSecret": "Web プラットフォームを使う場合のクライアントシークレット。public client の場合は空",
    "TenantId": "テナント ID",
    "GraphUserScopes": [
      "Sites.ReadWrite.All",
      "Files.ReadWrite.All"
    ],
    "TokenCachePath": ""
  }
}
```

SharePoint サイト ID は以下の形式で指定します。

```text
[テナント名].sharepoint.com,[Site GUID],[Web GUID]
```

GUID は SharePoint の REST API で確認できます。

```text
https://[テナント名].sharepoint.com/sites/[サイト名]/_api/site/id
https://[テナント名].sharepoint.com/sites/[サイト名]/_api/web/id
```

## 実行

```bash
dotnet run
```

初回実行後、ブラウザーで `/auth/login` を開いて Microsoft にサインインします。サインインが完了すると `/auth/callback` でトークンを受け取り、以後の MCP tool 呼び出しは保存されたトークンキャッシュを使います。

```text
http://localhost:3001/auth/login
```

認証状態は以下で確認できます。

```text
http://localhost:3001/auth/status
```

`TokenCachePath` を空にした場合、トークンキャッシュは OS のローカルアプリケーションデータ配下に保存されます。共有環境では、必要に応じてアクセス権を制限したパスを明示してください。

## 使用例

空の Excel ファイルを作成して編集セッションを開始します。

```csharp
var drive = await graphHelper.GetSiteDefaultDriveAsync(siteId);

var fileName = $"{Guid.NewGuid()}.xlsx";
var excelFile = await graphHelper.CreateEmptyExcelFile(drive, fileName);

await graphHelper.CreateExcelSession(drive, excelFile);
```

ワークシートを作成し、セルや Range に値を書き込みます。

```csharp
var worksheet = await graphHelper.CreateWorksheet(drive, excelFile, "売上");

await graphHelper.SetCellValue(drive, excelFile, worksheet, "B2", "設定した値");

await graphHelper.SetRangeValues(
    drive,
    excelFile,
    worksheet,
    "C3",
    new object?[][]
    {
        new object?[] { "商品", "金額" },
        new object?[] { "ノート", 1200 },
    });
```

有効な Range と、その Range 内の全セル値を取得します。

```csharp
var usedRange = await graphHelper.GetUsedRange(drive, excelFile, worksheet);
var values = await graphHelper.GetRangeValues(drive, excelFile, worksheet, usedRange);

foreach (var row in values)
{
    Console.WriteLine(string.Join(", ", row));
}
```

編集が終わったらセッションを閉じます。

```csharp
await graphHelper.CloseExcelSession(drive, excelFile);
```

## 補足

`WorkbookRange.Address` は `"Sheet1!A1:B2"` のようにシート名付きで返ることがあります。`GetRangeValues` の `WorkbookRange` オーバーロードでは、内部で `!` より後ろの `A1:B2` 部分を使って値を取得します。

## HTTP MCP サーバー

このアプリケーションは Streamable HTTP 対応の MCP サーバーとして起動します。

```bash
dotnet run --urls http://localhost:3001
```

MCP endpoint は以下です。

```text
http://localhost:3001/mcp
```

ブラウザーなどで `/` にアクセスすると、サーバー名、MCP endpoint、認証用 endpoint を返します。未認証の場合は `/auth/login` を開いて Microsoft にサインインしてから MCP tool を呼び出してください。

公開している MCP tool は以下です。

- `get_site_default_drive`
- `upload_file_base64`
- `download_file_base64`
- `create_empty_excel_file`
- `get_drive_item_by_path`
- `create_excel_session`
- `close_excel_session`
- `list_worksheets`
- `create_worksheet`
- `rename_worksheet`
- `set_cell_value`
- `get_cell_value`
- `set_range_values`
- `get_used_range`
- `get_range_values`
- `add_row`
- `add_rows`
- `insert_row`
- `insert_rows`
- `update_row`

各 tool は Graph SDK の `Drive` や `WorkbookWorksheet` オブジェクトを直接受け取らず、`siteId`、`filePath`、`worksheetName`、`cellAddress`、`rangeAddress`、`values` などの JSON 化しやすい引数で呼び出します。
