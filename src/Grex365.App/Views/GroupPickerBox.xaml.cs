using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.Views;

// Reusable group autocomplete box: type a name/mail and pick from live suggestions
// (IGroupsService.SearchAsync, debounced). Two-way Text DP holds the chosen group identifier —
// Mail when present, else DisplayName, which matches OnboardingService.ResolveGroupIdAsync's
// match order (Mail first, then DisplayName). Mirrors UserPickerBox.
public partial class GroupPickerBox : UserControl
{
    private readonly DispatcherTimer _debounce;
    private CancellationTokenSource? _cts;
    private bool _suppress;

    public GroupPickerBox()
    {
        InitializeComponent();
        _debounce = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
        _debounce.Tick += async (_, _) => { _debounce.Stop(); await SearchAsync(); };
    }

    public static readonly DependencyProperty TextProperty =
        DependencyProperty.Register(nameof(Text), typeof(string), typeof(GroupPickerBox),
            new FrameworkPropertyMetadata(string.Empty,
                FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnTextChanged));

    public string Text
    {
        get => (string)GetValue(TextProperty);
        set => SetValue(TextProperty, value);
    }

    public static readonly DependencyProperty PlaceholderProperty =
        DependencyProperty.Register(nameof(Placeholder), typeof(string), typeof(GroupPickerBox),
            new PropertyMetadata(string.Empty));

    public string Placeholder
    {
        get => (string)GetValue(PlaceholderProperty);
        set => SetValue(PlaceholderProperty, value);
    }

    private static void OnTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        var box = (GroupPickerBox)d;
        var value = e.NewValue as string ?? string.Empty;
        if (box.Input.Text != value)
        {
            box._suppress = true;
            box.Input.Text = value;
            box._suppress = false;
        }
    }

    private void Input_TextChanged(object sender, TextChangedEventArgs e)
    {
        Text = Input.Text;
        if (_suppress) return;
        _debounce.Stop();
        if (Input.Text.Trim().Length < 2)
        {
            Pop.IsOpen = false;
            return;
        }
        _debounce.Start();
    }

    private async Task SearchAsync()
    {
        var groups = App.Services?.GetService(typeof(IGroupsService)) as IGroupsService;
        if (groups is null) return;
        var query = Input.Text.Trim();
        if (query.Length < 2) return;

        _cts?.Cancel();
        _cts = new CancellationTokenSource();
        var token = _cts.Token;
        try
        {
            var found = await groups.SearchAsync(query, token).ConfigureAwait(true);
            if (token.IsCancellationRequested) return;
            var items = new List<GroupSummary>(found.Count > 12 ? 12 : found.Count);
            foreach (var g in found)
            {
                items.Add(g);
                if (items.Count >= 12) break;
            }
            List.ItemsSource = items;
            Pop.IsOpen = items.Count > 0 && Input.IsKeyboardFocusWithin;
        }
        catch
        {
            // typeahead stays quiet on errors
        }
    }

    private void List_Pick(object sender, MouseButtonEventArgs e) => Commit();

    private void List_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter) { Commit(); e.Handled = true; }
        else if (e.Key == Key.Escape) { Pop.IsOpen = false; e.Handled = true; }
    }

    private void Commit()
    {
        if (List.SelectedItem is not GroupSummary g) return;
        var value = !string.IsNullOrWhiteSpace(g.Mail) ? g.Mail! : g.DisplayName;
        if (string.IsNullOrWhiteSpace(value)) return;
        _suppress = true;
        Input.Text = value;
        Text = value;
        _suppress = false;
        Pop.IsOpen = false;
        Input.Focus();
    }

    private void Input_LostFocus(object sender, RoutedEventArgs e)
    {
        // Let a click on a suggestion process first, then close if focus really left.
        Dispatcher.BeginInvoke(new Action(() =>
        {
            if (!Pop.IsKeyboardFocusWithin && !Input.IsKeyboardFocusWithin)
            {
                Pop.IsOpen = false;
            }
        }), DispatcherPriority.Background);
    }
}
