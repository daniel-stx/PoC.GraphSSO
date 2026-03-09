using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;

namespace PoC.GraphSSO.Services;

public sealed class GraphEmployeeDirectoryService(
    GraphServiceClient graphServiceClient,
    ILogger<GraphEmployeeDirectoryService> logger) : IEmployeeDirectoryService
{
    #region User Creation PoC

    public async Task<UserCreateResult> CreateUserAsync(CreateUserRequest request, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.DisplayName))
        {
            return UserCreateResult.InvalidRequest("displayName is required.");
        }

        if (string.IsNullOrWhiteSpace(request.MailNickname))
        {
            return UserCreateResult.InvalidRequest("mailNickname is required.");
        }

        if (string.IsNullOrWhiteSpace(request.UserPrincipalName))
        {
            return UserCreateResult.InvalidRequest("userPrincipalName is required.");
        }

        if (string.IsNullOrWhiteSpace(request.Password))
        {
            return UserCreateResult.InvalidRequest("password is required.");
        }

        try
        {
            var createdUser = await graphServiceClient.Users.PostAsync(new User
            {
                AccountEnabled = request.AccountEnabled,
                DisplayName = request.DisplayName.Trim(),
                MailNickname = request.MailNickname.Trim(),
                UserPrincipalName = request.UserPrincipalName.Trim(),
                EmployeeId = string.IsNullOrWhiteSpace(request.EmployeeId) ? null : request.EmployeeId.Trim(),
                PasswordProfile = new PasswordProfile
                {
                    Password = request.Password,
                    ForceChangePasswordNextSignIn = request.ForceChangePasswordNextSignIn
                }
            }, cancellationToken: cancellationToken);

            if (createdUser is null)
            {
                return UserCreateResult.UnexpectedFailure("Microsoft Graph did not return the created user.");
            }

            logger.LogInformation("Created user {UserPrincipalName}.", createdUser.UserPrincipalName);

            return UserCreateResult.Success(
                createdUser.Id,
                createdUser.UserPrincipalName,
                createdUser.DisplayName,
                createdUser.EmployeeId,
                createdUser.AccountEnabled);
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 400)
        {
            logger.LogWarning(exception, "Microsoft Graph rejected create user request for {UserPrincipalName}.", request.UserPrincipalName);
            return UserCreateResult.InvalidRequest(
                "Microsoft Graph rejected the create user request. Verify the UPN domain, password policy, and required fields.");
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 403)
        {
            logger.LogWarning(exception, "Microsoft Graph denied create user request for {UserPrincipalName}.", request.UserPrincipalName);
            return UserCreateResult.PermissionDenied(
                "Microsoft Graph denied the create user request. Verify admin consent and User.ReadWrite.All application permission.");
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 409)
        {
            logger.LogWarning(exception, "Microsoft Graph reported a conflict while creating {UserPrincipalName}.", request.UserPrincipalName);
            return UserCreateResult.Conflict("A user with the same userPrincipalName or alias already exists.");
        }
        catch (ApiException exception)
        {
            logger.LogError(exception, "Unexpected Microsoft Graph error while creating {UserPrincipalName}.", request.UserPrincipalName);
            return UserCreateResult.UnexpectedFailure("Microsoft Graph returned an unexpected error while creating the user.");
        }
    }

    public async Task<GuestInvitationResult> CreateGuestInvitationAsync(
        GuestInvitationRequest request,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.InvitedUserEmailAddress))
        {
            return GuestInvitationResult.InvalidRequest("invitedUserEmailAddress is required.");
        }

        if (string.IsNullOrWhiteSpace(request.InviteRedirectUrl))
        {
            return GuestInvitationResult.InvalidRequest("inviteRedirectUrl is required.");
        }

        try
        {
            var invitation = await graphServiceClient.Invitations.PostAsync(new Invitation
            {
                InvitedUserEmailAddress = request.InvitedUserEmailAddress.Trim(),
                InviteRedirectUrl = request.InviteRedirectUrl.Trim(),
                SendInvitationMessage = request.SendInvitationMessage
            }, cancellationToken: cancellationToken);

            if (invitation is null)
            {
                return GuestInvitationResult.UnexpectedFailure("Microsoft Graph did not return the invitation.");
            }

            logger.LogInformation("Created guest invitation for {InvitedUserEmailAddress}.", invitation.InvitedUserEmailAddress);

            return GuestInvitationResult.Success(
                invitation.Id,
                invitation.InvitedUser?.Id,
                invitation.InvitedUserEmailAddress,
                invitation.InvitedUser?.UserPrincipalName,
                invitation.InviteRedeemUrl,
                invitation.Status);
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 400)
        {
            logger.LogWarning(exception, "Microsoft Graph rejected guest invitation for {InvitedUserEmailAddress}.", request.InvitedUserEmailAddress);
            return GuestInvitationResult.InvalidRequest(
                "Microsoft Graph rejected the guest invitation request. Verify the external email address and redirect URL.");
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 403)
        {
            logger.LogWarning(exception, "Microsoft Graph denied guest invitation for {InvitedUserEmailAddress}.", request.InvitedUserEmailAddress);
            return GuestInvitationResult.PermissionDenied(
                "Microsoft Graph denied the guest invitation request. Verify admin consent and guest invitation permissions.");
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 409)
        {
            logger.LogWarning(exception, "Microsoft Graph reported a conflict while inviting {InvitedUserEmailAddress}.", request.InvitedUserEmailAddress);
            return GuestInvitationResult.Conflict("A guest invitation conflict occurred for this external user.");
        }
        catch (ApiException exception)
        {
            logger.LogError(exception, "Unexpected Microsoft Graph error while inviting {InvitedUserEmailAddress}.", request.InvitedUserEmailAddress);
            return GuestInvitationResult.UnexpectedFailure("Microsoft Graph returned an unexpected error while creating the guest invitation.");
        }
    }

    public async Task<GuestInvitationResult> ReinviteGuestAsync(
        GuestReinviteRequest request,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.InvitedUserId))
        {
            return GuestInvitationResult.InvalidRequest("invitedUserId is required.");
        }

        if (string.IsNullOrWhiteSpace(request.InvitedUserEmailAddress))
        {
            return GuestInvitationResult.InvalidRequest("invitedUserEmailAddress is required.");
        }

        if (string.IsNullOrWhiteSpace(request.InviteRedirectUrl))
        {
            return GuestInvitationResult.InvalidRequest("inviteRedirectUrl is required.");
        }

        try
        {
            var invitation = await graphServiceClient.Invitations.PostAsync(new Invitation
            {
                InvitedUserEmailAddress = request.InvitedUserEmailAddress.Trim(),
                InviteRedirectUrl = request.InviteRedirectUrl.Trim(),
                SendInvitationMessage = request.SendInvitationMessage,
                ResetRedemption = true,
                InvitedUser = new User
                {
                    Id = request.InvitedUserId.Trim()
                }
            }, cancellationToken: cancellationToken);

            if (invitation is null)
            {
                return GuestInvitationResult.UnexpectedFailure("Microsoft Graph did not return the reinvitation.");
            }

            logger.LogInformation(
                "Triggered guest reinvitation for {InvitedUserEmailAddress} with user id {InvitedUserId}.",
                invitation.InvitedUserEmailAddress,
                request.InvitedUserId);

            return GuestInvitationResult.Success(
                invitation.Id,
                invitation.InvitedUser?.Id,
                invitation.InvitedUserEmailAddress,
                invitation.InvitedUser?.UserPrincipalName,
                invitation.InviteRedeemUrl,
                invitation.Status);
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 400)
        {
            logger.LogWarning(exception, "Microsoft Graph rejected guest reinvitation for {InvitedUserEmailAddress}.", request.InvitedUserEmailAddress);
            return GuestInvitationResult.InvalidRequest(
                "Microsoft Graph rejected the guest reinvitation request. Verify the guest user id, email address, and redirect URL.");
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 403)
        {
            logger.LogWarning(exception, "Microsoft Graph denied guest reinvitation for {InvitedUserEmailAddress}.", request.InvitedUserEmailAddress);
            return GuestInvitationResult.PermissionDenied(
                "Microsoft Graph denied the guest reinvitation request. Verify admin consent and guest invitation permissions.");
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 404)
        {
            logger.LogWarning(exception, "Microsoft Graph could not find guest user {InvitedUserId} for reinvitation.", request.InvitedUserId);
            return GuestInvitationResult.InvalidRequest("Microsoft Graph could not find the guest user for reinvitation.");
        }
        catch (ApiException exception)
        {
            logger.LogError(exception, "Unexpected Microsoft Graph error while reinviting {InvitedUserEmailAddress}.", request.InvitedUserEmailAddress);
            return GuestInvitationResult.UnexpectedFailure("Microsoft Graph returned an unexpected error while reinviting the guest user.");
        }
    }

    #endregion

    #region EmployeeId PoC

    public async Task<EmployeeIdQueryResult> GetEmployeeIdAsync(string userId, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(userId))
        {
            return EmployeeIdQueryResult.InvalidRequest("A target user id or user principal name is required.");
        }

        var trimmedUserId = userId.Trim();

        try
        {
            var user = await GetUserAsync(trimmedUserId, cancellationToken);

            if (user is null)
            {
                return EmployeeIdQueryResult.UserNotFound($"User '{trimmedUserId}' was not found.");
            }

            return EmployeeIdQueryResult.Success(user.Id ?? trimmedUserId, user.UserPrincipalName, user.EmployeeId);
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 404)
        {
            return EmployeeIdQueryResult.UserNotFound($"User '{trimmedUserId}' was not found.");
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 400)
        {
            logger.LogWarning(exception, "Microsoft Graph rejected employeeId lookup for user {UserId}.", trimmedUserId);
            return EmployeeIdQueryResult.InvalidRequest("Microsoft Graph rejected the employeeId lookup request.");
        }
        catch (ApiException exception)
        {
            logger.LogError(exception, "Unexpected Microsoft Graph error while reading employeeId for user {UserId}.", trimmedUserId);
            return EmployeeIdQueryResult.UnexpectedFailure("Microsoft Graph returned an unexpected error while reading employeeId.");
        }
    }

    public async Task<EmployeeIdUpdateResult> UpdateEmployeeIdAsync(
        string userId,
        string employeeId,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(userId))
        {
            return EmployeeIdUpdateResult.InvalidRequest("A target user id or user principal name is required.");
        }

        if (string.IsNullOrWhiteSpace(employeeId))
        {
            return EmployeeIdUpdateResult.InvalidRequest("employeeId is required.");
        }

        var trimmedUserId = userId.Trim();
        var trimmedEmployeeId = employeeId.Trim();

        try
        {
            var user = await GetUserAsync(trimmedUserId, cancellationToken);

            if (user is null)
            {
                return EmployeeIdUpdateResult.UserNotFound($"User '{trimmedUserId}' was not found.");
            }

            if (user.OnPremisesSyncEnabled == true)
            {
                return EmployeeIdUpdateResult.CloudManagedRequired(
                    $"User '{user.UserPrincipalName ?? user.Id ?? trimmedUserId}' is synchronized from on-premises, so employeeId must be managed at the source of authority.");
            }

            await graphServiceClient.Users[trimmedUserId].PatchAsync(new User
            {
                EmployeeId = trimmedEmployeeId
            }, cancellationToken: cancellationToken);

            logger.LogInformation("Updated employeeId for user {UserId}.", user.Id ?? trimmedUserId);

            return EmployeeIdUpdateResult.Success(user.Id ?? trimmedUserId, user.UserPrincipalName, trimmedEmployeeId);
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 404)
        {
            return EmployeeIdUpdateResult.UserNotFound($"User '{trimmedUserId}' was not found.");
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 403)
        {
            logger.LogWarning(exception, "Microsoft Graph denied employeeId update for user {UserId}.", trimmedUserId);
            return EmployeeIdUpdateResult.PermissionDenied(
                "Microsoft Graph denied the employeeId update. Verify admin consent and User.ReadWrite.All application permission.");
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 400)
        {
            logger.LogWarning(exception, "Microsoft Graph rejected employeeId update for user {UserId}.", trimmedUserId);
            return EmployeeIdUpdateResult.InvalidRequest(
                "Microsoft Graph rejected the employeeId update. Verify that the property is writable for this user and that the value is valid.");
        }
        catch (ApiException exception)
        {
            logger.LogError(exception, "Unexpected Microsoft Graph error while updating employeeId for user {UserId}.", trimmedUserId);
            return EmployeeIdUpdateResult.UnexpectedFailure(
                "Microsoft Graph returned an unexpected error while updating employeeId.");
        }
    }

    private Task<User?> GetUserAsync(string userId, CancellationToken cancellationToken) =>
        graphServiceClient.Users[userId].GetAsync(requestConfiguration =>
        {
            requestConfiguration.QueryParameters.Select =
            [
                "id",
                "userPrincipalName",
                "employeeId",
                "onPremisesSyncEnabled"
            ];
        }, cancellationToken);

    #endregion
}

#region User Creation PoC Models

public sealed record CreateUserRequest(
    string DisplayName,
    string MailNickname,
    string UserPrincipalName,
    string Password,
    bool AccountEnabled = true,
    bool ForceChangePasswordNextSignIn = true,
    string? EmployeeId = null);

public sealed record GuestInvitationRequest(
    string InvitedUserEmailAddress,
    string InviteRedirectUrl,
    bool SendInvitationMessage = true);

public sealed record GuestReinviteRequest(
    string InvitedUserId,
    string InvitedUserEmailAddress,
    string InviteRedirectUrl,
    bool SendInvitationMessage = true);

public enum UserCreateStatus
{
    Success,
    PermissionDenied,
    InvalidRequest,
    Conflict,
    UnexpectedFailure
}

public sealed record UserCreateResult(
    UserCreateStatus Status,
    string Message,
    string? UserId,
    string? UserPrincipalName,
    string? DisplayName,
    string? EmployeeId,
    bool? AccountEnabled)
{
    public static UserCreateResult Success(
        string? userId,
        string? userPrincipalName,
        string? displayName,
        string? employeeId,
        bool? accountEnabled) =>
        new(UserCreateStatus.Success, string.Empty, userId, userPrincipalName, displayName, employeeId, accountEnabled);

    public static UserCreateResult PermissionDenied(string message) =>
        new(UserCreateStatus.PermissionDenied, message, null, null, null, null, null);

    public static UserCreateResult InvalidRequest(string message) =>
        new(UserCreateStatus.InvalidRequest, message, null, null, null, null, null);

    public static UserCreateResult Conflict(string message) =>
        new(UserCreateStatus.Conflict, message, null, null, null, null, null);

    public static UserCreateResult UnexpectedFailure(string message) =>
        new(UserCreateStatus.UnexpectedFailure, message, null, null, null, null, null);
}

public enum GuestInvitationStatus
{
    Success,
    PermissionDenied,
    InvalidRequest,
    Conflict,
    UnexpectedFailure
}

public sealed record GuestInvitationResult(
    GuestInvitationStatus Status,
    string Message,
    string? InvitationId,
    string? InvitedUserId,
    string? InvitedUserEmailAddress,
    string? InvitedUserPrincipalName,
    string? InviteRedeemUrl,
    string? InvitationStatus)
{
    public static GuestInvitationResult Success(
        string? invitationId,
        string? invitedUserId,
        string? invitedUserEmailAddress,
        string? invitedUserPrincipalName,
        string? inviteRedeemUrl,
        string? invitationStatus) =>
        new(
            GuestInvitationStatus.Success,
            string.Empty,
            invitationId,
            invitedUserId,
            invitedUserEmailAddress,
            invitedUserPrincipalName,
            inviteRedeemUrl,
            invitationStatus);

    public static GuestInvitationResult PermissionDenied(string message) =>
        new(GuestInvitationStatus.PermissionDenied, message, null, null, null, null, null, null);

    public static GuestInvitationResult InvalidRequest(string message) =>
        new(GuestInvitationStatus.InvalidRequest, message, null, null, null, null, null, null);

    public static GuestInvitationResult Conflict(string message) =>
        new(GuestInvitationStatus.Conflict, message, null, null, null, null, null, null);

    public static GuestInvitationResult UnexpectedFailure(string message) =>
        new(GuestInvitationStatus.UnexpectedFailure, message, null, null, null, null, null, null);
}

#endregion

#region EmployeeId PoC Models

public enum EmployeeIdQueryStatus
{
    Success,
    UserNotFound,
    InvalidRequest,
    UnexpectedFailure
}

public sealed record EmployeeIdQueryResult(
    EmployeeIdQueryStatus Status,
    string Message,
    string? UserId,
    string? UserPrincipalName,
    string? EmployeeId)
{
    public static EmployeeIdQueryResult Success(string userId, string? userPrincipalName, string? employeeId) =>
        new(EmployeeIdQueryStatus.Success, string.Empty, userId, userPrincipalName, employeeId);

    public static EmployeeIdQueryResult UserNotFound(string message) =>
        new(EmployeeIdQueryStatus.UserNotFound, message, null, null, null);

    public static EmployeeIdQueryResult InvalidRequest(string message) =>
        new(EmployeeIdQueryStatus.InvalidRequest, message, null, null, null);

    public static EmployeeIdQueryResult UnexpectedFailure(string message) =>
        new(EmployeeIdQueryStatus.UnexpectedFailure, message, null, null, null);
}

public enum EmployeeIdUpdateStatus
{
    Success,
    UserNotFound,
    CloudManagedRequired,
    PermissionDenied,
    InvalidRequest,
    UnexpectedFailure
}

public sealed record EmployeeIdUpdateResult(
    EmployeeIdUpdateStatus Status,
    string Message,
    string? UserId,
    string? UserPrincipalName,
    string? EmployeeId)
{
    public static EmployeeIdUpdateResult Success(string userId, string? userPrincipalName, string? employeeId) =>
        new(EmployeeIdUpdateStatus.Success, string.Empty, userId, userPrincipalName, employeeId);

    public static EmployeeIdUpdateResult UserNotFound(string message) =>
        new(EmployeeIdUpdateStatus.UserNotFound, message, null, null, null);

    public static EmployeeIdUpdateResult CloudManagedRequired(string message) =>
        new(EmployeeIdUpdateStatus.CloudManagedRequired, message, null, null, null);

    public static EmployeeIdUpdateResult PermissionDenied(string message) =>
        new(EmployeeIdUpdateStatus.PermissionDenied, message, null, null, null);

    public static EmployeeIdUpdateResult InvalidRequest(string message) =>
        new(EmployeeIdUpdateStatus.InvalidRequest, message, null, null, null);

    public static EmployeeIdUpdateResult UnexpectedFailure(string message) =>
        new(EmployeeIdUpdateStatus.UnexpectedFailure, message, null, null, null);
}

#endregion
