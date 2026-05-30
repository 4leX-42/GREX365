using System.Windows;
using System.Windows.Controls;

namespace Grex365.App.Views;

public partial class UserDetailsView : UserControl
{
    public UserDetailsView()
    {
        InitializeComponent();
    }

    // When false the header Close button is hidden. The drawer (Groups) keeps it visible;
    // the inline Users panel sets ShowClose="False" since closing it makes no sense there.
    public static readonly DependencyProperty ShowCloseProperty =
        DependencyProperty.Register(
            nameof(ShowClose),
            typeof(bool),
            typeof(UserDetailsView),
            new PropertyMetadata(true));

    public bool ShowClose
    {
        get => (bool)GetValue(ShowCloseProperty);
        set => SetValue(ShowCloseProperty, value);
    }
}
