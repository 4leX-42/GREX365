using System;

namespace Grex365.App.Services;

public enum UpdateFeedKind
{
    None,    // empty / whitespace / invalid → updates disabled
    GitHub,  // a github.com repo URL → Velopack GithubSource (release assets)
    Http,    // any other absolute http(s) URL → simple directory feed
}

// Pure: classifies the configured update feed URL. Kept static + side-effect free so the
// routing is unit-testable without Velopack.
public static class UpdateFeedResolver
{
    public static UpdateFeedKind Classify(string? url)
    {
        var trimmed = url?.Trim();
        if (string.IsNullOrEmpty(trimmed)) return UpdateFeedKind.None;
        if (!Uri.TryCreate(trimmed, UriKind.Absolute, out var uri)) return UpdateFeedKind.None;
        if (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps) return UpdateFeedKind.None;
        return uri.Host.Equals("github.com", StringComparison.OrdinalIgnoreCase)
            || uri.Host.EndsWith(".github.com", StringComparison.OrdinalIgnoreCase)
            ? UpdateFeedKind.GitHub
            : UpdateFeedKind.Http;
    }
}
