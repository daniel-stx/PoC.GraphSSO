namespace PoC.GraphSSO.Services;

public interface IEmployeeDirectoryService
{
    #region User Creation PoC

    Task<UserCreateResult> CreateUserAsync(CreateUserRequest request, CancellationToken cancellationToken);

    Task<GuestInvitationResult> CreateGuestInvitationAsync(GuestInvitationRequest request, CancellationToken cancellationToken);

    Task<GuestInvitationResult> ReinviteGuestAsync(GuestReinviteRequest request, CancellationToken cancellationToken);

    #endregion

    #region EmployeeId PoC

    Task<EmployeeIdQueryResult> GetEmployeeIdAsync(string userId, CancellationToken cancellationToken);

    Task<EmployeeIdUpdateResult> UpdateEmployeeIdAsync(string userId, string employeeId, CancellationToken cancellationToken);

    #endregion
}
