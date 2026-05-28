using System.Runtime.CompilerServices;
using Grex365.App;
using Xunit;

// L10n is a process-wide static. ViewModels now resolve their StatusMessage / dialog
// text through L10n.Get at runtime, so tests asserting Spanish literals require the
// Spanish dictionary to be active. Two guarantees keep this race-free:
//   1. Collections run serially (the only L10n mutators live in the "L10n" collection,
//      but VM tests in other collections read L10n at runtime — serial avoids the race).
//   2. The module initializer below seeds Spanish before any test runs; the L10n-mutating
//      classes restore Spanish in their Dispose so every subsequent test sees the default.
[assembly: CollectionBehavior(DisableTestParallelization = true)]

namespace Grex365.App.Tests;

internal static class TestAssemblyConfig
{
    [ModuleInitializer]
    public static void Init() => L10n.Initialize("es");
}
