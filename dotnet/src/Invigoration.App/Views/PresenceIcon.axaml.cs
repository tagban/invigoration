using Avalonia;
using Avalonia.Controls;
using Invigoration.App.Models;

namespace Invigoration.App.Views;

public partial class PresenceIcon : UserControl
{
    public static readonly StyledProperty<PresenceState> StateProperty =
        AvaloniaProperty.Register<PresenceIcon, PresenceState>(nameof(State), PresenceState.Offline);

    public PresenceState State
    {
        get => GetValue(StateProperty);
        set => SetValue(StateProperty, value);
    }

    public PresenceIcon()
    {
        InitializeComponent();
        UpdateShape();
    }

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == StateProperty)
        {
            UpdateShape();
        }
    }

    private void UpdateShape()
    {
        SetVisible("AvailableShape", State == PresenceState.Available);
        SetVisible("AwayShape", State == PresenceState.Away);
        SetVisible("BusyShape", State == PresenceState.DoNotDisturb);
        SetVisible("InGameShape", State == PresenceState.InGame);
        SetVisible("OfflineShape", State == PresenceState.Offline);
    }

    private void SetVisible(string name, bool visible)
    {
        if (this.FindControl<Control>(name) is { } control)
        {
            control.IsVisible = visible;
        }
    }
}
