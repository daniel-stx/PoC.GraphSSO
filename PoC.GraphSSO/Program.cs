using System.Diagnostics;
using Azure.Identity;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using PoC.GraphSSO.Options;
using PoC.GraphSSO.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddOpenApi();

builder.Services.AddOptions<GraphApiOptions>()
    .Bind(builder.Configuration.GetSection(GraphApiOptions.SectionName))
    .ValidateDataAnnotations()
    .ValidateOnStart();

builder.Services.AddScoped(sp =>
{
    var options = sp.GetRequiredService<IOptions<GraphApiOptions>>().Value;
    var credential = new ClientSecretCredential(options.TenantId, options.ClientId, options.ClientSecret);

    // App-only tokens keep employeeId updates independent from the signed-in user's Graph privileges.
    return new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);
});
builder.Services.AddScoped<IEmployeeDirectoryService, GraphEmployeeDirectoryService>();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

app.UseHttpsRedirection();

app.MapGet("/", () => Results.Ok(new
    {
        Message = "Graph-only PoC for reading and updating employeeId.",
        Endpoints = new[]
        {
            "GET /users/{userId}/employee-id",
            "POST /users/{userId}/employee-id"
        }
    }))
    .WithName("Home");

app.MapGet("/users/{userId}/employee-id",
        async (string userId, IEmployeeDirectoryService employeeDirectoryService, CancellationToken cancellationToken) =>
        {
            var stopwatch = Stopwatch.StartNew();
            var result = await employeeDirectoryService.GetEmployeeIdAsync(userId, cancellationToken);
            stopwatch.Stop();

            return result.Status switch
            {
                EmployeeIdQueryStatus.Success => Results.Ok(new
                {
                    userId = result.UserId,
                    userPrincipalName = result.UserPrincipalName,
                    employeeId = result.EmployeeId,
                    durationMs = stopwatch.ElapsedMilliseconds
                }),
                EmployeeIdQueryStatus.UserNotFound => Results.NotFound(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                EmployeeIdQueryStatus.InvalidRequest => Results.BadRequest(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                _ => Results.Problem(statusCode: StatusCodes.Status502BadGateway, detail: result.Message)
            };
        })
    .WithName("GetEmployeeId");

app.MapPost("/users/{userId}/employee-id",
        async (string userId, UpdateEmployeeIdRequest request, IEmployeeDirectoryService employeeDirectoryService,
            CancellationToken cancellationToken) =>
        {
            var stopwatch = Stopwatch.StartNew();
            var result = await employeeDirectoryService.UpdateEmployeeIdAsync(userId, request.EmployeeId, cancellationToken);
            stopwatch.Stop();

            return result.Status switch
            {
                EmployeeIdUpdateStatus.Success => Results.Ok(new
                {
                    userId = result.UserId,
                    userPrincipalName = result.UserPrincipalName,
                    employeeId = result.EmployeeId,
                    durationMs = stopwatch.ElapsedMilliseconds
                }),
                EmployeeIdUpdateStatus.UserNotFound => Results.NotFound(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                EmployeeIdUpdateStatus.CloudManagedRequired => Results.Conflict(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                EmployeeIdUpdateStatus.PermissionDenied => Results.Json(
                    new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds },
                    statusCode: StatusCodes.Status403Forbidden),
                EmployeeIdUpdateStatus.InvalidRequest => Results.BadRequest(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                _ => Results.Problem(statusCode: StatusCodes.Status502BadGateway, detail: result.Message)
            };
        })
    .WithName("UpdateEmployeeId");

app.Run();

internal sealed record UpdateEmployeeIdRequest(string EmployeeId);