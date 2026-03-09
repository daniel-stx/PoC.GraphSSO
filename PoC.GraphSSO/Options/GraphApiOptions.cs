using System.ComponentModel.DataAnnotations;

namespace PoC.GraphSSO.Options;

public sealed class GraphApiOptions
{
    public const string SectionName = "Graph";

    [Required]
    public string TenantId { get; set; } = string.Empty;

    [Required]
    public string ClientId { get; set; } = string.Empty;

    [Required]
    public string ClientSecret { get; set; } = string.Empty;
}
