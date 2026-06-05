using System;
using System.Threading;
using System.Threading.Tasks;
using Grex365.Core.Abstractions;
using Velopack;
using Velopack.Sources;

namespace Grex365.App.Services;

// Velopack UpdateManager wrapper. The feed URL lives in preferences (UpdateFeedUrl) so it can
// point at the public binaries-only GitHub repo or an internal HTTP share without recompiling.
public sealed class VelopackUpdateService : IUpdateService
{
    private readonly IPreferencesStore _prefs;
    private UpdateInfo? _pending;
    private UpdateManager? _manager;

    public VelopackUpdateService(IPreferencesStore prefs)
    {
        _prefs = prefs;
    }

    public async Task<UpdateCheckResult> CheckAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            var url = (await _prefs.LoadAsync(cancellationToken).ConfigureAwait(false)).UpdateFeedUrl?.Trim();
            var kind = UpdateFeedResolver.Classify(url);
            if (kind == UpdateFeedKind.None)
            {
                return new UpdateCheckResult(UpdateCheckStatus.FeedNotConfigured);
            }

            var manager = kind switch
            {
                UpdateFeedKind.GitHub => new UpdateManager(new GithubSource(url!, accessToken: null, prerelease: false)),
                UpdateFeedKind.LocalPath => new UpdateManager(new SimpleFileSource(
                    new System.IO.DirectoryInfo(UpdateFeedResolver.ToLocalDirectory(url!)))),
                _ => new UpdateManager(new SimpleWebSource(url!)),
            };

            if (!manager.IsInstalled)
            {
                // Dev build or portable copy outside a Velopack install — nothing to update.
                return new UpdateCheckResult(UpdateCheckStatus.NotInstalled);
            }

            var info = await manager.CheckForUpdatesAsync().ConfigureAwait(false);
            if (info is null)
            {
                _pending = null;
                _manager = null;
                return new UpdateCheckResult(UpdateCheckStatus.UpToDate);
            }

            _pending = info;
            _manager = manager;
            return new UpdateCheckResult(UpdateCheckStatus.UpdateAvailable, info.TargetFullRelease.Version.ToString());
        }
        catch (Exception ex)
        {
            return new UpdateCheckResult(UpdateCheckStatus.Error, Detail: ex.Message);
        }
    }

    public async Task ApplyAsync(CancellationToken cancellationToken = default)
    {
        var manager = _manager;
        var pending = _pending;
        if (manager is null || pending is null)
        {
            throw new InvalidOperationException("No hay actualización pendiente. Ejecuta primero la comprobación.");
        }
        await manager.DownloadUpdatesAsync(pending, cancelToken: cancellationToken).ConfigureAwait(false);
        manager.ApplyUpdatesAndRestart(pending);
    }
}
