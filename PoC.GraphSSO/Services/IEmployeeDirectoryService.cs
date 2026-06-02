namespace PoC.GraphSSO.Services;

public interface IEmployeeDirectoryService
{
    #region User Creation PoC

    Task<UserCreateResult> CreateUserAsync(CreateUserRequest request, CancellationToken cancellationToken);

    Task<GuestInvitationResult> CreateGuestInvitationAsync(GuestInvitationRequest request, CancellationToken cancellationToken);

    Task<GuestInvitationResult> ReinviteGuestAsync(GuestReinviteRequest request, CancellationToken cancellationToken);

    #endregion

    #region User Properties PoC

    Task<UserPropertiesQueryResult> GetUserPropertiesAsync(string userId, CancellationToken cancellationToken);

    Task<UserPropertiesUpdateResult> UpdateUserPropertiesAsync(
        string userId,
        UserPropertiesUpdateRequest request,
        CancellationToken cancellationToken);

    #endregion

    #region Group Membership PoC

    Task<GroupMembershipResult> AddUserToGroupAsync(string groupId, string userId, CancellationToken cancellationToken);

    #endregion
}
