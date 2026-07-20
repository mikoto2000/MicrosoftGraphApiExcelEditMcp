using edit_excel;

using Microsoft.AspNetCore.HttpOverrides;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddSingleton(Settings.LoadSettings());
builder.Services.AddSingleton(serviceProvider =>
{
  var settings = serviceProvider.GetRequiredService<Settings>();
  var logger = serviceProvider.GetRequiredService<ILogger<Program>>();

  return GraphHelper.InitializeGraphForUserAuth(
    settings,
    (info, cancel) =>
    {
      logger.LogWarning("Device code authentication required: {Message}", info.Message);
      return Task.CompletedTask;
    });
});

builder.Services
  .AddMcpServer()
  .WithHttpTransport(options =>
  {
    options.Stateless = true;
  })
  .WithToolsFromAssembly();

var app = builder.Build();

app.UseForwardedHeaders(new ForwardedHeadersOptions
{
  ForwardedHeaders = ForwardedHeaders.XForwardedFor | ForwardedHeaders.XForwardedProto,
});

app.MapGet("/", () => Results.Ok(new
{
  name = "Microsoft Graph Excel MCP Server",
  endpoint = "/mcp",
}));
app.MapMcp("/mcp");

app.Run();
