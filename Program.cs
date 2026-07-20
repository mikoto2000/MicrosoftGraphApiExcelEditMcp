using edit_excel;

using System.Security.Claims;

using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.IdentityModel.Tokens;
using Microsoft.AspNetCore.HttpOverrides;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddSingleton(Settings.LoadSettings());
builder.Services.AddSingleton(serviceProvider =>
{
  var settings = serviceProvider.GetRequiredService<Settings>();
  return GraphHelper.InitializeGraphForUserAuth(settings);
});

builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
  .AddJwtBearer(options =>
  {
    var settings = Settings.LoadSettings();
    var tenantId = settings.TenantId ?? throw new NullReferenceException("Settings.TenantId cannot be null");
    var clientId = settings.ClientId ?? throw new NullReferenceException("Settings.ClientId cannot be null");
    options.Authority = $"https://login.microsoftonline.com/{tenantId}/v2.0";
    options.TokenValidationParameters = new TokenValidationParameters
    {
      ValidateIssuer = true,
      ValidIssuers = new[]
      {
        $"https://login.microsoftonline.com/{tenantId}/v2.0",
        $"https://sts.windows.net/{tenantId}/",
      },
      ValidAudiences = new[]
      {
        settings.ApiAudience,
        $"api://{clientId}",
        clientId,
      }.Where(value => !string.IsNullOrWhiteSpace(value)),
    };
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

app.UseAuthentication();

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
  var bearerToken = GetValidatedBearerToken(context);
  if (HasBearerToken(context) && bearerToken == null)
  {
    context.Response.StatusCode = StatusCodes.Status401Unauthorized;
    await context.Response.WriteAsync("Invalid bearer token for this MCP server.");
    return;
  }

  using (app.Services.GetRequiredService<GraphHelper>().UseRequestContext(GetProfile(context), bearerToken))
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

  if (string.IsNullOrWhiteSpace(profile) && context.User.Identity?.IsAuthenticated == true)
  {
    var tenantId = context.User.FindFirstValue("tid") ?? "unknown-tenant";
    var objectId = context.User.FindFirstValue("oid") ?? context.User.FindFirstValue(ClaimTypes.NameIdentifier) ?? "unknown-user";
    profile = $"{tenantId}.{objectId}";
  }

  return GraphHelper.NormalizeProfile(profile);
}

static bool HasBearerToken(HttpContext context)
{
  return context.Request.Headers.Authorization.ToString().StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase);
}

static string? GetValidatedBearerToken(HttpContext context)
{
  if (context.User.Identity?.IsAuthenticated != true || !HasBearerToken(context))
  {
    return null;
  }

  return context.Request.Headers.Authorization.ToString()["Bearer ".Length..].Trim();
}
