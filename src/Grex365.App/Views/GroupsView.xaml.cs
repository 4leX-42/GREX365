using System.Windows;
using System.Windows.Controls;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Extensions.DependencyInjection;

namespace Grex365.App.Views;

public partial class GroupsView : UserControl
{
    public GroupsView()
    {
        InitializeComponent();
    }

    private void MemberItem_DoubleClick(object sender, RoutedEventArgs e)
    {
        if (sender is not ListBoxItem item || item.DataContext is not GroupMember member)
        {
            return;
        }
        if (string.IsNullOrEmpty(member.Id))
        {
            return;
        }
        var host = App.Services.GetService<IUserDetailsHost>();
        host?.RequestOpen(member.Id);
    }
}
