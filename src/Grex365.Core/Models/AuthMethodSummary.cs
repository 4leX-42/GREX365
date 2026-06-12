namespace Grex365.Core.Models;

// One registered authentication method (MFA) on a user. Kind is a stable app-level tag
// (Phone / Fido2 / MicrosoftAuthenticator / WindowsHello / Email / SoftwareOath /
// TemporaryAccessPass / Password / Other). Password can never be deleted via Graph and
// unknown kinds have no typed delete endpoint — both carry Removable=false.
public sealed record AuthMethodSummary(
    string Id,
    string Kind,
    string? Detail,
    bool Removable);
