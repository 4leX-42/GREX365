using System.Collections.Concurrent;
using System.ComponentModel;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.Tests;

internal sealed class TestDialogService : IDialogService
{
    public bool ConfirmResult { get; set; } = true;
    public List<(string Message, string Title, DialogIcon Icon)> Confirmations { get; } = new();
    public List<(string Message, string Title, DialogIcon Icon)> Shows { get; } = new();

    public Task<bool> ConfirmAsync(string message, string title, DialogIcon icon = DialogIcon.Question, CancellationToken cancellationToken = default)
    {
        Confirmations.Add((message, title, icon));
        return Task.FromResult(ConfirmResult);
    }

    public Task ShowAsync(string message, string title, DialogIcon icon = DialogIcon.Info, CancellationToken cancellationToken = default)
    {
        Shows.Add((message, title, icon));
        return Task.CompletedTask;
    }
}

internal sealed class TestClipboardService : IClipboardService
{
    public string? LastValue { get; private set; }
    public int CallCount { get; private set; }
    public void SetText(string text)
    {
        LastValue = text;
        CallCount++;
    }
}

internal sealed class TestUiLogSink : IUiLogSink
{
    public System.Collections.ObjectModel.ObservableCollection<LogEntry> Entries { get; } = new();
    public IProgress<LogEntry> Progress { get; }

    public TestUiLogSink()
    {
        Progress = new TestProgress(this);
    }

    public void Clear() => Entries.Clear();

    private sealed class TestProgress : IProgress<LogEntry>
    {
        private readonly TestUiLogSink _owner;
        public TestProgress(TestUiLogSink owner) => _owner = owner;
        public void Report(LogEntry value) => _owner.Entries.Add(value);
    }
}

internal sealed class TestUserDetailsHost : IUserDetailsHost
{
    public bool IsOpen { get; private set; }
    public string? CurrentUserId { get; private set; }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event EventHandler<string>? OpenRequested;
    public event EventHandler? CloseRequested;

    public void RequestOpen(string userId)
    {
        CurrentUserId = userId;
        IsOpen = true;
        OpenRequested?.Invoke(this, userId);
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsOpen)));
    }

    public void RequestClose()
    {
        IsOpen = false;
        CurrentUserId = null;
        CloseRequested?.Invoke(this, EventArgs.Empty);
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsOpen)));
    }
}
