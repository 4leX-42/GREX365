using Grex365.Core.Models;

namespace Grex365.App.Services;

public interface INotifier
{
    void Notify(string title, string message, LogSeverity severity);
}
