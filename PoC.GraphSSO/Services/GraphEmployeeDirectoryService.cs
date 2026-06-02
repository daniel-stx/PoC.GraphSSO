using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;

namespace PoC.GraphSSO.Services;

public sealed class GraphEmployeeDirectoryService(
    GraphServiceClient graphServiceClient,
    ILogger<GraphEmployeeDirectoryService> logger) : IEmployeeDirectoryService
{
    private const string SynXisUsernameExtensionPropertyName = "extension_789456df0b4f43b19bef3d896030cb99_SynXisUsername";

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

        if (request.SynXisUsername is not null && string.IsNullOrWhiteSpace(request.SynXisUsername))
        {
            return UserCreateResult.InvalidRequest("synXisUsername cannot be empty.");
        }

        try
        {
            var trimmedSynXisUsername = request.SynXisUsername?.Trim();

            var createdUser = await graphServiceClient.Users.PostAsync(new User
            {
                AccountEnabled = request.AccountEnabled,
                DisplayName = request.DisplayName.Trim(),
                MailNickname = request.MailNickname.Trim(),
                UserPrincipalName = request.UserPrincipalName.Trim(),
                EmployeeId = string.IsNullOrWhiteSpace(request.EmployeeId) ? null : request.EmployeeId.Trim(),
                AdditionalData = trimmedSynXisUsername is null
                    ? null
                    : new Dictionary<string, object>
                    {
                        [SynXisUsernameExtensionPropertyName] = trimmedSynXisUsername
                    },
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
                trimmedSynXisUsername,
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

    #region User Properties PoC

    public async Task<UserPropertiesQueryResult> GetUserPropertiesAsync(string userId, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(userId))
        {
            return UserPropertiesQueryResult.InvalidRequest("A target user id or user principal name is required.");
        }

        var trimmedUserId = userId.Trim();

        try
        {
            var user = await GetUserAsync(trimmedUserId, cancellationToken);
            if (user is null)
            {
                return UserPropertiesQueryResult.UserNotFound($"User '{trimmedUserId}' was not found.");
            }

            return UserPropertiesQueryResult.Success(
                user.Id ?? trimmedUserId,
                user.UserPrincipalName,
                user.EmployeeId,
                GetSynXisUsername(user));
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 404)
        {
            return UserPropertiesQueryResult.UserNotFound($"User '{trimmedUserId}' was not found.");
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 400)
        {
            logger.LogWarning(exception, "Microsoft Graph rejected user properties lookup for user {UserId}.", trimmedUserId);
            return UserPropertiesQueryResult.InvalidRequest("Microsoft Graph rejected the user properties lookup request.");
        }
        catch (ApiException exception)
        {
            logger.LogError(exception, "Unexpected Microsoft Graph error while reading user properties for user {UserId}.", trimmedUserId);
            return UserPropertiesQueryResult.UnexpectedFailure("Microsoft Graph returned an unexpected error while reading user properties.");
        }
    }

    public async Task<UserPropertiesUpdateResult> UpdateUserPropertiesAsync(
        string userId,
        UserPropertiesUpdateRequest request,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(userId))
        {
            return UserPropertiesUpdateResult.InvalidRequest("A target user id or user principal name is required.");
        }

        if (request.EmployeeId is not null && string.IsNullOrWhiteSpace(request.EmployeeId))
        {
            return UserPropertiesUpdateResult.InvalidRequest("employeeId cannot be empty.");
        }

        if (request.SynXisUsername is not null && string.IsNullOrWhiteSpace(request.SynXisUsername))
        {
            return UserPropertiesUpdateResult.InvalidRequest("synXisUsername cannot be empty.");
        }

        if (request.EmployeeId is null && request.SynXisUsername is null)
        {
            return UserPropertiesUpdateResult.InvalidRequest("At least one user property is required.");
        }

        var trimmedUserId = userId.Trim();
        var trimmedEmployeeId = request.EmployeeId?.Trim();
        var trimmedSynXisUsername = request.SynXisUsername?.Trim();

        try
        {
            var user = await GetUserAsync(trimmedUserId, cancellationToken);

            if (user is null)
            {
                return UserPropertiesUpdateResult.UserNotFound($"User '{trimmedUserId}' was not found.");
            }

            if (trimmedEmployeeId is not null && user.OnPremisesSyncEnabled == true)
            {
                return UserPropertiesUpdateResult.CloudManagedRequired(
                    $"User '{user.UserPrincipalName ?? user.Id ?? trimmedUserId}' is synchronized from on-premises, so employeeId must be managed at the source of authority.");
            }

            var update = new User();

            if (trimmedEmployeeId is not null)
            {
                update.EmployeeId = trimmedEmployeeId;
            }

            if (trimmedSynXisUsername is not null)
            {
                update.AdditionalData = new Dictionary<string, object>
                {
                    [SynXisUsernameExtensionPropertyName] = trimmedSynXisUsername
                };
            }

            await graphServiceClient.Users[trimmedUserId].PatchAsync(update, cancellationToken: cancellationToken);

            logger.LogInformation("Updated user properties for user {UserId}.", user.Id ?? trimmedUserId);

            return UserPropertiesUpdateResult.Success(
                user.Id ?? trimmedUserId,
                user.UserPrincipalName,
                trimmedEmployeeId ?? user.EmployeeId,
                trimmedSynXisUsername ?? GetSynXisUsername(user));
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 404)
        {
            return UserPropertiesUpdateResult.UserNotFound($"User '{trimmedUserId}' was not found.");
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 403)
        {
            logger.LogWarning(exception, "Microsoft Graph denied user properties update for user {UserId}.", trimmedUserId);
            return UserPropertiesUpdateResult.PermissionDenied(
                "Microsoft Graph denied the user properties update. Verify admin consent and User.ReadWrite.All application permission.");
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 400)
        {
            logger.LogWarning(exception, "Microsoft Graph rejected user properties update for user {UserId}.", trimmedUserId);
            return UserPropertiesUpdateResult.InvalidRequest(
                "Microsoft Graph rejected the user properties update. Verify that the properties are writable for this user and that the values are valid.");
        }
        catch (ApiException exception)
        {
            logger.LogError(exception, "Unexpected Microsoft Graph error while updating user properties for user {UserId}.", trimmedUserId);
            return UserPropertiesUpdateResult.UnexpectedFailure(
                "Microsoft Graph returned an unexpected error while updating user properties.");
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
                "onPremisesSyncEnabled",
                SynXisUsernameExtensionPropertyName
            ];
        }, cancellationToken);

    private static string? GetSynXisUsername(User user) =>
        user.AdditionalData?.TryGetValue(SynXisUsernameExtensionPropertyName, out var value) == true
            ? value?.ToString()
            : null;

    #endregion

    #region Group Membership PoC

    public async Task<GroupMembershipResult> AddUserToGroupAsync(
        string groupId,
        string userId,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(groupId))
        {
            return GroupMembershipResult.InvalidRequest("SynXis SSO group id is not configured.");
        }

        if (string.IsNullOrWhiteSpace(userId))
        {
            return GroupMembershipResult.InvalidRequest("A target user object id is required.");
        }

        var trimmedGroupId = groupId.Trim();
        var trimmedUserId = userId.Trim();

        try
        {
            await graphServiceClient.Groups[trimmedGroupId].Members.Ref.PostAsync(new ReferenceCreate
            {
                OdataId = $"https://graph.microsoft.com/v1.0/directoryObjects/{trimmedUserId}"
            }, cancellationToken: cancellationToken);

            logger.LogInformation("Added user {UserId} to group {GroupId}.", trimmedUserId, trimmedGroupId);

            return GroupMembershipResult.Success(trimmedGroupId, trimmedUserId);
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 400 && IsAlreadyMemberError(exception))
        {
            logger.LogInformation("User {UserId} is already a member of group {GroupId}.", trimmedUserId, trimmedGroupId);
            return GroupMembershipResult.AlreadyMember(trimmedGroupId, trimmedUserId);
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 400)
        {
            logger.LogWarning(exception, "Microsoft Graph rejected group membership request for user {UserId} and group {GroupId}.", trimmedUserId, trimmedGroupId);
            return GroupMembershipResult.InvalidRequest(
                "Microsoft Graph rejected the group membership request. Verify that groupId and userId are valid object ids.");
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 403)
        {
            logger.LogWarning(exception, "Microsoft Graph denied group membership request for user {UserId} and group {GroupId}.", trimmedUserId, trimmedGroupId);
            return GroupMembershipResult.PermissionDenied(
                "Microsoft Graph denied the group membership request. Verify admin consent and GroupMember.ReadWrite.All application permission.");
        }
        catch (ApiException exception) when (exception.ResponseStatusCode == 404)
        {
            logger.LogWarning(exception, "Microsoft Graph could not find user {UserId} or group {GroupId}.", trimmedUserId, trimmedGroupId);
            return GroupMembershipResult.NotFound("Microsoft Graph could not find the target user or SynXis SSO group.");
        }
        catch (ApiException exception)
        {
            logger.LogError(exception, "Unexpected Microsoft Graph error while adding user {UserId} to group {GroupId}.", trimmedUserId, trimmedGroupId);
            return GroupMembershipResult.UnexpectedFailure("Microsoft Graph returned an unexpected error while adding the user to the group.");
        }
    }

    private static bool IsAlreadyMemberError(ApiException exception) =>
        exception.Message.Contains("already exist", StringComparison.OrdinalIgnoreCase)
        || exception.Message.Contains("already a member", StringComparison.OrdinalIgnoreCase);

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
    string? EmployeeId = null,
    string? SynXisUsername = null);

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
    string? SynXisUsername,
    bool? AccountEnabled)
{
    public static UserCreateResult Success(
        string? userId,
        string? userPrincipalName,
        string? displayName,
        string? employeeId,
        string? synXisUsername,
        bool? accountEnabled) =>
        new(UserCreateStatus.Success, string.Empty, userId, userPrincipalName, displayName, employeeId, synXisUsername, accountEnabled);

    public static UserCreateResult PermissionDenied(string message) =>
        new(UserCreateStatus.PermissionDenied, message, null, null, null, null, null, null);

    public static UserCreateResult InvalidRequest(string message) =>
        new(UserCreateStatus.InvalidRequest, message, null, null, null, null, null, null);

    public static UserCreateResult Conflict(string message) =>
        new(UserCreateStatus.Conflict, message, null, null, null, null, null, null);

    public static UserCreateResult UnexpectedFailure(string message) =>
        new(UserCreateStatus.UnexpectedFailure, message, null, null, null, null, null, null);
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

#region User Properties PoC Models

public sealed record UserPropertiesUpdateRequest(string? EmployeeId, string? SynXisUsername);

public enum UserPropertiesQueryStatus
{
    Success,
    UserNotFound,
    InvalidRequest,
    UnexpectedFailure
}

public sealed record UserPropertiesQueryResult(
    UserPropertiesQueryStatus Status,
    string Message,
    string? UserId,
    string? UserPrincipalName,
    string? EmployeeId,
    string? SynXisUsername)
{
    public static UserPropertiesQueryResult Success(
        string userId,
        string? userPrincipalName,
        string? employeeId,
        string? synXisUsername) =>
        new(UserPropertiesQueryStatus.Success, string.Empty, userId, userPrincipalName, employeeId, synXisUsername);

    public static UserPropertiesQueryResult UserNotFound(string message) =>
        new(UserPropertiesQueryStatus.UserNotFound, message, null, null, null, null);

    public static UserPropertiesQueryResult InvalidRequest(string message) =>
        new(UserPropertiesQueryStatus.InvalidRequest, message, null, null, null, null);

    public static UserPropertiesQueryResult UnexpectedFailure(string message) =>
        new(UserPropertiesQueryStatus.UnexpectedFailure, message, null, null, null, null);
}

public enum UserPropertiesUpdateStatus
{
    Success,
    UserNotFound,
    CloudManagedRequired,
    PermissionDenied,
    InvalidRequest,
    UnexpectedFailure
}

public sealed record UserPropertiesUpdateResult(
    UserPropertiesUpdateStatus Status,
    string Message,
    string? UserId,
    string? UserPrincipalName,
    string? EmployeeId,
    string? SynXisUsername)
{
    public static UserPropertiesUpdateResult Success(
        string userId,
        string? userPrincipalName,
        string? employeeId,
        string? synXisUsername) =>
        new(UserPropertiesUpdateStatus.Success, string.Empty, userId, userPrincipalName, employeeId, synXisUsername);

    public static UserPropertiesUpdateResult UserNotFound(string message) =>
        new(UserPropertiesUpdateStatus.UserNotFound, message, null, null, null, null);

    public static UserPropertiesUpdateResult CloudManagedRequired(string message) =>
        new(UserPropertiesUpdateStatus.CloudManagedRequired, message, null, null, null, null);

    public static UserPropertiesUpdateResult PermissionDenied(string message) =>
        new(UserPropertiesUpdateStatus.PermissionDenied, message, null, null, null, null);

    public static UserPropertiesUpdateResult InvalidRequest(string message) =>
        new(UserPropertiesUpdateStatus.InvalidRequest, message, null, null, null, null);

    public static UserPropertiesUpdateResult UnexpectedFailure(string message) =>
        new(UserPropertiesUpdateStatus.UnexpectedFailure, message, null, null, null, null);
}

#endregion

#region Group Membership PoC Models

public enum GroupMembershipStatus
{
    Success,
    AlreadyMember,
    NotFound,
    PermissionDenied,
    InvalidRequest,
    UnexpectedFailure
}

public sealed record GroupMembershipResult(
    GroupMembershipStatus Status,
    string Message,
    string? GroupId,
    string? UserId)
{
    public static GroupMembershipResult Success(string groupId, string userId) =>
        new(GroupMembershipStatus.Success, string.Empty, groupId, userId);

    public static GroupMembershipResult AlreadyMember(string groupId, string userId) =>
        new(GroupMembershipStatus.AlreadyMember, string.Empty, groupId, userId);

    public static GroupMembershipResult NotFound(string message) =>
        new(GroupMembershipStatus.NotFound, message, null, null);

    public static GroupMembershipResult PermissionDenied(string message) =>
        new(GroupMembershipStatus.PermissionDenied, message, null, null);

    public static GroupMembershipResult InvalidRequest(string message) =>
        new(GroupMembershipStatus.InvalidRequest, message, null, null);

    public static GroupMembershipResult UnexpectedFailure(string message) =>
        new(GroupMembershipStatus.UnexpectedFailure, message, null, null);
}

#endregion
