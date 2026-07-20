using edit_excel;

using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.AspNetCore.HttpOverrides;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddSingleton(Settings.LoadSettings());
builder.Services.AddSingleton(serviceProvider =>
{
  var settings = serviceProvider.GetRequiredService<Settings>();
  return GraphHelper.InitializeGraphForUserAuth(settings);
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
  profile = new
  {
    header = "X-Excel-Mcp-Profile",
    query = "profile",
    defaultValue = GraphHelper.DefaultProfile,
  },
  authentication = new
  {
    login = "/auth/login?profile=default",
    callback = "/auth/callback",
    status = "/auth/status?profile=default",
  },
}));

app.MapGet("/auth/status", async (HttpContext context, GraphHelper graph) =>
{
  var profile = GetProfile(context);
  return Results.Ok(new
  {
    profile,
    authenticated = await graph.IsAuthenticatedAsync(profile),
  });
});

app.MapGet("/auth/login", async (HttpContext context, GraphHelper graph) =>
{
  var profile = GetProfile(context);
  var redirectUri = BuildAuthCallbackUri(context);
  var authorizationUrl = graph.GetAuthorizationUrl(redirectUri, profile);
  return Results.Redirect(authorizationUrl.ToString());
});

app.MapGet("/auth/callback", async (HttpContext context, GraphHelper graph) =>
{
  if (context.Request.Query.TryGetValue("error", out var error))
  {
    var description = context.Request.Query.TryGetValue("error_description", out var errorDescription)
      ? errorDescription.ToString()
      : string.Empty;
    return Results.BadRequest(new { error = error.ToString(), error_description = description });
  }

  var code = context.Request.Query["code"].ToString();
  var state = context.Request.Query["state"].ToString();
  if (string.IsNullOrWhiteSpace(code))
  {
    return Results.BadRequest(new { error = "missing_code" });
  }

  try
  {
    var profile = await graph.CompleteAuthorizationCodeFlowAsync(code, state, BuildAuthCallbackUri(context));
    return Results.Content($"Microsoft Graph authentication completed for profile '{profile}'. You can close this page and retry the MCP tool call.", "text/plain");
  }
  catch (Exception ex)
  {
    return Results.Problem(ex.Message, statusCode: StatusCodes.Status400BadRequest);
  }
});

app.Use(async (context, next) =>
{
  using (app.Services.GetRequiredService<GraphHelper>().UseProfile(GetProfile(context)))
  {
    await next(context);
  }
});

app.MapMcp("/mcp");

app.Run();


static string BuildAuthCallbackUri(HttpContext context)
{
  return UriHelper.BuildAbsolute(
    context.Request.Scheme,
    context.Request.Host,
    context.Request.PathBase,
    "/auth/callback");
}


static string GetProfile(HttpContext context)
{
  var profile = context.Request.Query["profile"].ToString();
  if (string.IsNullOrWhiteSpace(profile) && context.Request.Headers.TryGetValue("X-Excel-Mcp-Profile", out var headerValues))
  {
    profile = headerValues.ToString();
  }

  return GraphHelper.NormalizeProfile(profile);
}
