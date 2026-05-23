namespace Grex365.App.ViewModels;

public static class NavRequirements
{
    // True iff the item's connection requirements are satisfied by the current state.
    // Item requires Graph → graphConnected must be true. Item requires EXO → exchangeConnected must be true.
    // Item with no requirements is always enabled.
    public static bool IsEnabled(bool requiresGraph, bool requiresExchange, bool graphConnected, bool exchangeConnected)
    {
        return (!requiresGraph || graphConnected)
            && (!requiresExchange || exchangeConnected);
    }

    public static bool IsEnabled(NavigationItem item, bool graphConnected, bool exchangeConnected)
    {
        ArgumentNullException.ThrowIfNull(item);
        return IsEnabled(item.RequiresGraph, item.RequiresExchange, graphConnected, exchangeConnected);
    }
}
