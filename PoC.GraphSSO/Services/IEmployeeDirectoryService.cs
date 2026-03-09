namespace PoC.GraphSSO.Services;

public interface IEmployeeDirectoryService
{
    Task<EmployeeIdQueryResult> GetEmployeeIdAsync(string userId, CancellationToken cancellationToken);

    Task<EmployeeIdUpdateResult> UpdateEmployeeIdAsync(string userId, string employeeId, CancellationToken cancellationToken);
}
