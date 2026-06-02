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

builder.Services.AddOptions<SynXisOptions>()
    .Bind(builder.Configuration.GetSection(SynXisOptions.SectionName));

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
var userPropertiesPoc = app.MapGroup("/poc/user-properties");
var synXisPoc = app.MapGroup("/poc/synxis");

#endregion

#region Home

app.MapGet("/", () => Results.Ok(new
    {
        Message = "Graph-only PoCs for user creation and user properties update.",
        Endpoints = new[]
        {
            "POST /poc/user-creation/users",
            "POST /poc/user-creation/invitations",
            "POST /poc/user-creation/invitations/reinvite",
            "GET /poc/user-properties/users/{userId}",
            "POST /poc/user-properties/users/{userId}",
            "POST /poc/synxis/users/{userId}/enable-sso",
            "POST /poc/synxis/users/{userId}/disable-sso"
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
                    request.EmployeeId,
                    request.SynXisUsername),
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
                    synXisUsername = result.SynXisUsername,
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

#region User Properties PoC

userPropertiesPoc.MapGet("/users/{userId}",
        async (string userId, IEmployeeDirectoryService employeeDirectoryService, CancellationToken cancellationToken) =>
        {
            var stopwatch = Stopwatch.StartNew();
            var result = await employeeDirectoryService.GetUserPropertiesAsync(userId, cancellationToken);
            stopwatch.Stop();

            return result.Status switch
            {
                UserPropertiesQueryStatus.Success => Results.Ok(new
                {
                    userId = result.UserId,
                    userPrincipalName = result.UserPrincipalName,
                    employeeId = result.EmployeeId,
                    synXisUsername = result.SynXisUsername,
                    durationMs = stopwatch.ElapsedMilliseconds
                }),
                UserPropertiesQueryStatus.UserNotFound => Results.NotFound(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                UserPropertiesQueryStatus.InvalidRequest => Results.BadRequest(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                _ => Results.Problem(statusCode: StatusCodes.Status502BadGateway, detail: result.Message)
            };
        })
    .WithName("GetUserProperties");

userPropertiesPoc.MapPost("/users/{userId}",
        async (string userId, UpdateUserPropertiesApiRequest request, IEmployeeDirectoryService employeeDirectoryService,
            CancellationToken cancellationToken) =>
        {
            var stopwatch = Stopwatch.StartNew();
            var result = await employeeDirectoryService.UpdateUserPropertiesAsync(
                userId,
                new UserPropertiesUpdateRequest(request.EmployeeId, request.SynXisUsername),
                cancellationToken);
            stopwatch.Stop();

            return result.Status switch
            {
                UserPropertiesUpdateStatus.Success => Results.Ok(new
                {
                    userId = result.UserId,
                    userPrincipalName = result.UserPrincipalName,
                    employeeId = result.EmployeeId,
                    synXisUsername = result.SynXisUsername,
                    durationMs = stopwatch.ElapsedMilliseconds
                }),
                UserPropertiesUpdateStatus.UserNotFound => Results.NotFound(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                UserPropertiesUpdateStatus.CloudManagedRequired => Results.Conflict(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                UserPropertiesUpdateStatus.PermissionDenied => Results.Json(
                    new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds },
                    statusCode: StatusCodes.Status403Forbidden),
                UserPropertiesUpdateStatus.InvalidRequest => Results.BadRequest(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                _ => Results.Problem(statusCode: StatusCodes.Status502BadGateway, detail: result.Message)
            };
        })
    .WithName("UpdateUserProperties");

#endregion

#region SynXis SSO PoC

synXisPoc.MapPost("/users/{userId}/enable-sso",
        async (string userId, IOptions<SynXisOptions> synXisOptions, IEmployeeDirectoryService employeeDirectoryService,
            CancellationToken cancellationToken) =>
        {
            var stopwatch = Stopwatch.StartNew();
            var result = await employeeDirectoryService.AddUserToGroupAsync(
                synXisOptions.Value.SsoGroupId,
                userId,
                cancellationToken);
            stopwatch.Stop();

            return result.Status switch
            {
                GroupMembershipStatus.Success => Results.Ok(new
                {
                    status = "Added",
                    groupId = result.GroupId,
                    userId = result.UserId,
                    durationMs = stopwatch.ElapsedMilliseconds
                }),
                GroupMembershipStatus.AlreadyMember => Results.Ok(new
                {
                    status = "AlreadyMember",
                    groupId = result.GroupId,
                    userId = result.UserId,
                    durationMs = stopwatch.ElapsedMilliseconds
                }),
                GroupMembershipStatus.NotFound => Results.NotFound(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                GroupMembershipStatus.PermissionDenied => Results.Json(
                    new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds },
                    statusCode: StatusCodes.Status403Forbidden),
                GroupMembershipStatus.InvalidRequest => Results.BadRequest(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                _ => Results.Problem(statusCode: StatusCodes.Status502BadGateway, detail: result.Message)
            };
        })
    .WithName("EnableSynXisSso");

synXisPoc.MapPost("/users/{userId}/disable-sso",
        async (string userId, IOptions<SynXisOptions> synXisOptions, IEmployeeDirectoryService employeeDirectoryService,
            CancellationToken cancellationToken) =>
        {
            var stopwatch = Stopwatch.StartNew();
            var result = await employeeDirectoryService.RemoveUserFromGroupAsync(
                synXisOptions.Value.SsoGroupId,
                userId,
                cancellationToken);
            stopwatch.Stop();

            return result.Status switch
            {
                GroupMembershipStatus.Success => Results.Ok(new
                {
                    status = "Removed",
                    groupId = result.GroupId,
                    userId = result.UserId,
                    durationMs = stopwatch.ElapsedMilliseconds
                }),
                GroupMembershipStatus.NotMember => Results.Ok(new
                {
                    status = "NotMember",
                    groupId = result.GroupId,
                    userId = result.UserId,
                    durationMs = stopwatch.ElapsedMilliseconds
                }),
                GroupMembershipStatus.NotFound => Results.NotFound(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                GroupMembershipStatus.PermissionDenied => Results.Json(
                    new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds },
                    statusCode: StatusCodes.Status403Forbidden),
                GroupMembershipStatus.InvalidRequest => Results.BadRequest(new { error = result.Message, durationMs = stopwatch.ElapsedMilliseconds }),
                _ => Results.Problem(statusCode: StatusCodes.Status502BadGateway, detail: result.Message)
            };
        })
    .WithName("DisableSynXisSso");

#endregion

app.Run();

internal sealed record CreateUserApiRequest(
    string DisplayName,
    string MailNickname,
    string UserPrincipalName,
    string Password,
    bool AccountEnabled = true,
    bool ForceChangePasswordNextSignIn = true,
    string? EmployeeId = null,
    string? SynXisUsername = null);

internal sealed record CreateGuestInvitationApiRequest(
    string InvitedUserEmailAddress,
    string InviteRedirectUrl,
    bool SendInvitationMessage = true);

internal sealed record ReinviteGuestApiRequest(
    string InvitedUserId,
    string InvitedUserEmailAddress,
    string InviteRedirectUrl,
    bool SendInvitationMessage = true);

internal sealed record UpdateUserPropertiesApiRequest(string? EmployeeId, string? SynXisUsername);