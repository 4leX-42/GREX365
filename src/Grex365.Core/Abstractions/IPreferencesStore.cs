using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface IPreferencesStore
{
    Task<UserPreferences> LoadAsync(CancellationToken cancellationToken = default);

    Task SaveAsync(UserPreferences preferences, CancellationToken cancellationToken = default);
}

public interface ICertConfigStore
{
    Task<CertConfig?> LoadAsync(CancellationToken cancellationToken = default);

    Task SaveAsync(CertConfig config, CancellationToken cancellationToken = default);

    Task DeleteAsync(CancellationToken cancellationToken = default);
}
