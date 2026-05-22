using System.Windows;
using System.Windows.Controls;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Extensions.DependencyInjection;

namespace Grex365.App.Views;

public partial class UsersView : UserControl
{
    public UsersView()
    {
        InitializeComponent();
    }

    private void UserItem_DoubleClick(object sender, RoutedEventArgs e)
    {
        if (sender is not ListBoxItem item || item.DataContext is not UserSummary user)
        {
            return;
        }
        if (string.IsNullOrEmpty(user.Id))
        {
            return;
        }
        var host = App.Services.GetService<IUserDetailsHost>();
        host?.RequestOpen(user.Id);
    }
}
