using System.ComponentModel;
using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface IConnectionStateMonitor : INotifyPropertyChanged, IAsyncDisposable
{
    ConnectionState Current { get; }

    void Start();

    void Stop();
}
