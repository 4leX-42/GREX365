using System;
using Velopack;

namespace Grex365.App;

// Custom entry point: Velopack must run FIRST (it handles install/update/uninstall hooks and
// may exit the process during those events). Declared as StartupObject in the csproj so this
// Main wins over the WPF-generated one.
public static class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        VelopackApp.Build().Run();

        var app = new App();
        app.InitializeComponent();
        app.Run();
    }
}
