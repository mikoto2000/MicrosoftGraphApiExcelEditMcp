using System.ComponentModel;
using System.Text.Json;

using Microsoft.Graph.Models;
using ModelContextProtocol.Server;

namespace edit_excel;

[McpServerToolType]
public sealed class GraphHelperTools(GraphHelper graph)
{
  [McpServerTool(Name = "get_site_default_drive", ReadOnly = true), Description("Gets the default document drive for a SharePoint site.")]
  public async Task<DriveSummary> GetSiteDefaultDrive(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId)
  {
    var drive = await graph.GetSiteDefaultDriveAsync(siteId);
    return ToDriveSummary(drive);
  }

  [McpServerTool(Name = "upload_file_base64"), Description("Uploads a base64-encoded file to the root of the site's default drive.")]
  public async Task<DriveItemSummary> UploadFileBase64(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Destination file name or root-relative path in the drive.")] string filePath,
    [Description("Base64-encoded file content.")] string base64Content)
  {
    var drive = await graph.GetSiteDefaultDriveAsync(siteId);
    using var stream = new MemoryStream(Convert.FromBase64String(base64Content));
    await graph.UploadFile(drive, filePath, stream);
    return ToDriveItemSummary(await graph.GetDriveItemByPath(drive, filePath));
  }

  [McpServerTool(Name = "download_file_base64", ReadOnly = true), Description("Downloads a file from the root of the site's default drive as base64.")]
  public async Task<Base64FileResult> DownloadFileBase64(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("File name or root-relative path in the drive.")] string filePath)
  {
    var drive = await graph.GetSiteDefaultDriveAsync(siteId);
    await using var stream = await graph.DownloadFile(drive, filePath);
    using var memory = new MemoryStream();
    await stream.CopyToAsync(memory);
    return new Base64FileResult(filePath, Convert.ToBase64String(memory.ToArray()));
  }

  [McpServerTool(Name = "create_empty_excel_file"), Description("Creates an empty .xlsx file in the root of the site's default drive.")]
  public async Task<DriveItemSummary> CreateEmptyExcelFile(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path. Must end with .xlsx.")] string filePath)
  {
    var drive = await graph.GetSiteDefaultDriveAsync(siteId);
    return ToDriveItemSummary(await graph.CreateEmptyExcelFile(drive, filePath));
  }

  [McpServerTool(Name = "get_drive_item_by_path", ReadOnly = true), Description("Gets a drive item from the root of the site's default drive by path.")]
  public async Task<DriveItemSummary> GetDriveItemByPath(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("File name or root-relative path in the drive.")] string filePath)
  {
    var (drive, file) = await ResolveFile(siteId, filePath);
    return ToDriveItemSummary(file);
  }

  [McpServerTool(Name = "create_excel_session"), Description("Creates a persistent Excel workbook session for the specified file.")]
  public async Task<SessionResult> CreateExcelSession(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path in the drive.")] string filePath)
  {
    var (drive, file) = await ResolveFile(siteId, filePath);
    await graph.CreateExcelSession(drive, file);
    return new SessionResult(file.Id, true);
  }

  [McpServerTool(Name = "close_excel_session"), Description("Closes the Excel workbook session for the specified file if one exists.")]
  public async Task<SessionResult> CloseExcelSession(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path in the drive.")] string filePath)
  {
    var (drive, file) = await ResolveFile(siteId, filePath);
    await graph.CloseExcelSession(drive, file);
    return new SessionResult(file.Id, false);
  }

  [McpServerTool(Name = "list_worksheets", ReadOnly = true), Description("Lists worksheets in an Excel workbook.")]
  public async Task<IReadOnlyList<WorksheetSummary>> ListWorksheets(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path in the drive.")] string filePath)
  {
    var (drive, file) = await ResolveFile(siteId, filePath);
    var worksheets = await graph.CountExcelSheet(drive, file);
    return worksheets.Select(ToWorksheetSummary).ToArray();
  }

  [McpServerTool(Name = "create_worksheet"), Description("Creates a worksheet in an Excel workbook.")]
  public async Task<WorksheetSummary> CreateWorksheet(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path in the drive.")] string filePath,
    [Description("Name for the new worksheet.")] string worksheetName)
  {
    var (drive, file) = await ResolveFile(siteId, filePath);
    return ToWorksheetSummary(await graph.CreateWorksheet(drive, file, worksheetName));
  }

  [McpServerTool(Name = "rename_worksheet"), Description("Renames a worksheet in an Excel workbook.")]
  public async Task<WorksheetSummary> RenameWorksheet(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path in the drive.")] string filePath,
    [Description("Current worksheet name.")] string worksheetName,
    [Description("New worksheet name.")] string newName)
  {
    var (drive, file, worksheet) = await ResolveWorksheet(siteId, filePath, worksheetName);
    return ToWorksheetSummary(await graph.RenameWorksheet(drive, file, worksheet, newName));
  }

  [McpServerTool(Name = "set_cell_value"), Description("Sets a single cell value in a worksheet.")]
  public async Task<RangeSummary> SetCellValue(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path in the drive.")] string filePath,
    [Description("Worksheet name.")] string worksheetName,
    [Description("A1-style cell address, such as B2.")] string cellAddress,
    [Description("JSON value to set in the cell.")] JsonElement value)
  {
    var (drive, file, worksheet) = await ResolveWorksheet(siteId, filePath, worksheetName);
    return ToRangeSummary(await graph.SetCellValue(drive, file, worksheet, cellAddress, FromJsonElement(value)));
  }

  [McpServerTool(Name = "get_cell_value", ReadOnly = true), Description("Gets a single cell value from a worksheet.")]
  public async Task<object?> GetCellValue(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path in the drive.")] string filePath,
    [Description("Worksheet name.")] string worksheetName,
    [Description("A1-style cell address, such as B2.")] string cellAddress)
  {
    var (drive, file, worksheet) = await ResolveWorksheet(siteId, filePath, worksheetName);
    return await graph.GetCellValue(drive, file, worksheet, cellAddress);
  }

  [McpServerTool(Name = "set_range_values"), Description("Sets values in a worksheet range, using start cell plus a two-dimensional JSON array.")]
  public async Task<RangeSummary> SetRangeValues(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path in the drive.")] string filePath,
    [Description("Worksheet name.")] string worksheetName,
    [Description("A1-style start cell address, such as C3.")] string startCellAddress,
    [Description("Two-dimensional JSON array of cell values.")] JsonElement[][] values)
  {
    var (drive, file, worksheet) = await ResolveWorksheet(siteId, filePath, worksheetName);
    return ToRangeSummary(await graph.SetRangeValues(drive, file, worksheet, startCellAddress, ToObjectRows(values)));
  }

  [McpServerTool(Name = "get_used_range", ReadOnly = true), Description("Gets the used range metadata for a worksheet.")]
  public async Task<RangeSummary> GetUsedRange(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path in the drive.")] string filePath,
    [Description("Worksheet name.")] string worksheetName)
  {
    var (drive, file, worksheet) = await ResolveWorksheet(siteId, filePath, worksheetName);
    return ToRangeSummary(await graph.GetUsedRange(drive, file, worksheet));
  }

  [McpServerTool(Name = "get_range_values", ReadOnly = true), Description("Gets all values from a worksheet range.")]
  public async Task<object?[][]> GetRangeValues(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path in the drive.")] string filePath,
    [Description("Worksheet name.")] string worksheetName,
    [Description("A1-style range address, such as A1:C10.")] string rangeAddress)
  {
    var (drive, file, worksheet) = await ResolveWorksheet(siteId, filePath, worksheetName);
    return await graph.GetRangeValues(drive, file, worksheet, rangeAddress);
  }

  [McpServerTool(Name = "add_row"), Description("Adds one row after the worksheet's used range.")]
  public async Task<RangeSummary> AddRow(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path in the drive.")] string filePath,
    [Description("Worksheet name.")] string worksheetName,
    [Description("JSON array of cell values.")] JsonElement[] row)
  {
    var (drive, file, worksheet) = await ResolveWorksheet(siteId, filePath, worksheetName);
    return ToRangeSummary(await graph.AddRow(drive, file, worksheet, ToNonNullObjectRow(row)));
  }

  [McpServerTool(Name = "add_rows"), Description("Adds rows after the worksheet's used range.")]
  public async Task<RangeSummary> AddRows(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path in the drive.")] string filePath,
    [Description("Worksheet name.")] string worksheetName,
    [Description("Two-dimensional JSON array of row values.")] JsonElement[][] rows)
  {
    var (drive, file, worksheet) = await ResolveWorksheet(siteId, filePath, worksheetName);
    return ToRangeSummary(await graph.AddRows(drive, file, worksheet, ToObjectRows(rows).Select(row => row.Select(cell => cell ?? string.Empty).ToArray()).ToArray()));
  }

  [McpServerTool(Name = "insert_row"), Description("Inserts one row at the specified 1-based Excel row number.")]
  public async Task<RangeSummary> InsertRow(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path in the drive.")] string filePath,
    [Description("Worksheet name.")] string worksheetName,
    [Description("1-based Excel row number.")] int rowNumber,
    [Description("JSON array of cell values.")] JsonElement[] row)
  {
    var (drive, file, worksheet) = await ResolveWorksheet(siteId, filePath, worksheetName);
    return ToRangeSummary(await graph.InsertRow(drive, file, worksheet, rowNumber, ToNonNullObjectRow(row)));
  }

  [McpServerTool(Name = "insert_rows"), Description("Inserts rows at the specified 1-based Excel row number.")]
  public async Task<RangeSummary> InsertRows(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path in the drive.")] string filePath,
    [Description("Worksheet name.")] string worksheetName,
    [Description("1-based Excel row number.")] int rowNumber,
    [Description("Two-dimensional JSON array of row values.")] JsonElement[][] rows)
  {
    var (drive, file, worksheet) = await ResolveWorksheet(siteId, filePath, worksheetName);
    return ToRangeSummary(await graph.InsertRows(drive, file, worksheet, rowNumber, ToObjectRows(rows).Select(row => row.Select(cell => cell ?? string.Empty).ToArray()).ToArray()));
  }

  [McpServerTool(Name = "update_row"), Description("Updates one row at the specified 1-based Excel row number.")]
  public async Task<RangeSummary> UpdateRow(
    [Description("SharePoint site ID in the form tenant.sharepoint.com,site-guid,web-guid.")] string siteId,
    [Description("Excel file name or root-relative path in the drive.")] string filePath,
    [Description("Worksheet name.")] string worksheetName,
    [Description("1-based Excel row number.")] int rowNumber,
    [Description("JSON array of cell values.")] JsonElement[] row)
  {
    var (drive, file, worksheet) = await ResolveWorksheet(siteId, filePath, worksheetName);
    return ToRangeSummary(await graph.UpdateRow(drive, file, worksheet, rowNumber, ToNonNullObjectRow(row)));
  }

  private async Task<(Drive Drive, DriveItem File)> ResolveFile(string siteId, string filePath)
  {
    var drive = await graph.GetSiteDefaultDriveAsync(siteId);
    var file = await graph.GetDriveItemByPath(drive, filePath);
    return (drive, file);
  }

  private async Task<(Drive Drive, DriveItem File, WorkbookWorksheet Worksheet)> ResolveWorksheet(string siteId, string filePath, string worksheetName)
  {
    var (drive, file) = await ResolveFile(siteId, filePath);
    var worksheets = await graph.CountExcelSheet(drive, file);
    var worksheet = worksheets.FirstOrDefault(item => string.Equals(item.Name, worksheetName, StringComparison.OrdinalIgnoreCase));
    if (worksheet == null)
    {
      throw new ArgumentException($"Worksheet '{worksheetName}' was not found.", nameof(worksheetName));
    }

    return (drive, file, worksheet);
  }

  private static object?[][] ToObjectRows(JsonElement[][] values) => values.Select(ToObjectRow).ToArray();

  private static object?[] ToObjectRow(JsonElement[] row) => row.Select(FromJsonElement).ToArray();

  private static object[] ToNonNullObjectRow(JsonElement[] row) => ToObjectRow(row).Select(cell => cell ?? string.Empty).ToArray();

  private static object? FromJsonElement(JsonElement element) => element.ValueKind switch
  {
    JsonValueKind.Null => null,
    JsonValueKind.Undefined => null,
    JsonValueKind.String => element.GetString(),
    JsonValueKind.True => true,
    JsonValueKind.False => false,
    JsonValueKind.Number when element.TryGetInt32(out var intValue) => intValue,
    JsonValueKind.Number when element.TryGetInt64(out var longValue) => longValue,
    JsonValueKind.Number when element.TryGetDouble(out var doubleValue) => doubleValue,
    JsonValueKind.Array => element.EnumerateArray().Select(FromJsonElement).ToArray(),
    JsonValueKind.Object => element.EnumerateObject().ToDictionary(property => property.Name, property => FromJsonElement(property.Value)),
    _ => element.ToString(),
  };

  private static DriveSummary ToDriveSummary(Drive drive) => new(drive.Id, drive.Name, drive.WebUrl);

  private static DriveItemSummary ToDriveItemSummary(DriveItem item) => new(item.Id, item.Name, item.WebUrl, item.Size);

  private static WorksheetSummary ToWorksheetSummary(WorkbookWorksheet worksheet) => new(worksheet.Id, worksheet.Name, worksheet.Position, worksheet.Visibility);

  private static RangeSummary ToRangeSummary(WorkbookRange range) => new(range.Address, range.RowIndex, range.ColumnIndex, range.RowCount, range.ColumnCount);
}

public sealed record DriveSummary(string? Id, string? Name, string? WebUrl);

public sealed record DriveItemSummary(string? Id, string? Name, string? WebUrl, long? Size);

public sealed record WorksheetSummary(string? Id, string? Name, int? Position, string? Visibility);

public sealed record RangeSummary(string? Address, int? RowIndex, int? ColumnIndex, int? RowCount, int? ColumnCount);

public sealed record Base64FileResult(string FilePath, string Base64Content);

public sealed record SessionResult(string? DriveItemId, bool IsOpen);
