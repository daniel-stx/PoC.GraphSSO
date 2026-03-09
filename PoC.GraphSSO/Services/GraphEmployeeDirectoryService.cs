using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;

namespace PoC.GraphSSO.Services;

public sealed class GraphEmployeeDirectoryService(
    GraphServiceClient graphServiceClient,
    ILogger<GraphEmployeeDirectoryService> logger) : IEmployeeDirectoryService
{
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

            var xdd = await graphServiceClient.Users[trimmedUserId].PatchAsync(new User
            {
                EmployeeId = trimmedEmployeeId
            }, cancellationToken: cancellationToken);

            var test = await GetUserAsync(trimmedUserId, cancellationToken);

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
}

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
