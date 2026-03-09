using System.Diagnostics;
using Azure.Identity;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using PoC.GraphSSO.Options;
using PoC.GraphSSO.Services;

#region Bootstrap

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

#endregion

#region Route Groups

var userCreationPoc = app.MapGroup("/poc/user-creation");
var employeeIdPoc = app.MapGroup("/poc/employee-id");

#endregion

#region Home

app.MapGet("/", () => Results.Ok(new
    {
        Message = "Graph-only PoCs for user creation and employeeId update.",
        Endpoints = new[]
        {
            "POST /poc/user-creation/users",
            "POST /poc/user-creation/invitations",
            "POST /poc/user-creation/invitations/reinvite",
            "GET /poc/employee-id/users/{userId}/employee-id",
            "POST /poc/employee-id/users/{userId}/employee-id"
        }
    }))
    .WithName("Home");

#endregion

#region User Creation PoC

userCreationPoc.MapPost("/users",
        async (CreateUserApiRequest request, IEmployeeDirectoryService employeeDirectoryService,
            CancellationToken cancellationToken) =>
        {
            var stopwatch = Stopwatch.StartNew();
            var result = await employeeDirectoryService.CreateUserAsync(
                new CreateUserRequest(
                    request.DisplayName,
                    request.MailNickname,
                    request.UserPrincipalName,
                    request.Password,
                    request.AccountEnabled,
                    request.ForceChangePasswordNextSignIn,
                    request.EmployeeId),
                cancellationToken);
            stopwatch.Stop();

            return result.Status switch
            {
                UserCreateStatus.Success => Results.Ok(new
                {
                    userId = result.UserId,
                    userPrincipalName = result.UserPrincipalName,
                    displayName = result.DisplayName,
                    employeeId = result.EmployeeId,
                    accountEnabled = result.AccountEnabled,
                    durationMs = stopwatch.ElapsedMilliseconds
                }),
                UserCreateStatus.Conflict => Results.Conflict(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                UserCreateStatus.PermissionDenied => Results.Json(
                    new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds },
                    statusCode: StatusCodes.Status403Forbidden),
                UserCreateStatus.InvalidRequest => Results.BadRequest(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                _ => Results.Problem(statusCode: StatusCodes.Status502BadGateway, detail: result.Message)
            };
        })
    .WithName("CreateUser");

userCreationPoc.MapPost("/invitations",
        async (CreateGuestInvitationApiRequest request, IEmployeeDirectoryService employeeDirectoryService,
            CancellationToken cancellationToken) =>
        {
            var stopwatch = Stopwatch.StartNew();
            var result = await employeeDirectoryService.CreateGuestInvitationAsync(
                new GuestInvitationRequest(
                    request.InvitedUserEmailAddress,
                    request.InviteRedirectUrl,
                    request.SendInvitationMessage),
                cancellationToken);
            stopwatch.Stop();

            return result.Status switch
            {
                GuestInvitationStatus.Success => Results.Ok(new
                {
                    invitationId = result.InvitationId,
                    invitedUserId = result.InvitedUserId,
                    invitedUserEmailAddress = result.InvitedUserEmailAddress,
                    invitedUserPrincipalName = result.InvitedUserPrincipalName,
                    inviteRedeemUrl = result.InviteRedeemUrl,
                    invitationStatus = result.InvitationStatus,
                    durationMs = stopwatch.ElapsedMilliseconds
                }),
                GuestInvitationStatus.Conflict => Results.Conflict(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                GuestInvitationStatus.PermissionDenied => Results.Json(
                    new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds },
                    statusCode: StatusCodes.Status403Forbidden),
                GuestInvitationStatus.InvalidRequest => Results.BadRequest(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                _ => Results.Problem(statusCode: StatusCodes.Status502BadGateway, detail: result.Message)
            };
        })
    .WithName("CreateGuestInvitation");

userCreationPoc.MapPost("/invitations/reinvite",
        async (ReinviteGuestApiRequest request, IEmployeeDirectoryService employeeDirectoryService,
            CancellationToken cancellationToken) =>
        {
            var stopwatch = Stopwatch.StartNew();
            var result = await employeeDirectoryService.ReinviteGuestAsync(
                new GuestReinviteRequest(
                    request.InvitedUserId,
                    request.InvitedUserEmailAddress,
                    request.InviteRedirectUrl,
                    request.SendInvitationMessage),
                cancellationToken);
            stopwatch.Stop();

            return result.Status switch
            {
                GuestInvitationStatus.Success => Results.Ok(new
                {
                    invitationId = result.InvitationId,
                    invitedUserId = result.InvitedUserId,
                    invitedUserEmailAddress = result.InvitedUserEmailAddress,
                    invitedUserPrincipalName = result.InvitedUserPrincipalName,
                    inviteRedeemUrl = result.InviteRedeemUrl,
                    invitationStatus = result.InvitationStatus,
                    durationMs = stopwatch.ElapsedMilliseconds
                }),
                GuestInvitationStatus.Conflict => Results.Conflict(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                GuestInvitationStatus.PermissionDenied => Results.Json(
                    new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds },
                    statusCode: StatusCodes.Status403Forbidden),
                GuestInvitationStatus.InvalidRequest => Results.BadRequest(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                _ => Results.Problem(statusCode: StatusCodes.Status502BadGateway, detail: result.Message)
            };
        })
    .WithName("ReinviteGuest");

#endregion

#region EmployeeId PoC

employeeIdPoc.MapGet("/users/{userId}/employee-id",
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

employeeIdPoc.MapPost("/users/{userId}/employee-id",
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

#endregion

app.Run();

internal sealed record CreateUserApiRequest(
    string DisplayName,
    string MailNickname,
    string UserPrincipalName,
    string Password,
    bool AccountEnabled = true,
    bool ForceChangePasswordNextSignIn = true,
    string? EmployeeId = null);

internal sealed record CreateGuestInvitationApiRequest(
    string InvitedUserEmailAddress,
    string InviteRedirectUrl,
    bool SendInvitationMessage = true);

internal sealed record ReinviteGuestApiRequest(
    string InvitedUserId,
    string InvitedUserEmailAddress,
    string InviteRedirectUrl,
    bool SendInvitationMessage = true);

internal sealed record UpdateEmployeeIdRequest(string EmployeeId);