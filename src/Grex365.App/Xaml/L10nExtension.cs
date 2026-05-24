using System.Windows.Markup;

namespace Grex365.App.Xaml;

[MarkupExtensionReturnType(typeof(string))]
public sealed class L10nExtension : MarkupExtension
{
    public string? Key { get; set; }

    public L10nExtension() { }

    public L10nExtension(string key) => Key = key;

    public override object ProvideValue(IServiceProvider serviceProvider)
    {
        return string.IsNullOrEmpty(Key) ? string.Empty : L10n.Get(Key);
    }
}
