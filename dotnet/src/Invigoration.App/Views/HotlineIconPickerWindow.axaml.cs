using Avalonia.Controls;
using Avalonia.Interactivity;
using Invigoration.App.ViewModels;

namespace Invigoration.App.Views;

/// <summary>ShowDialog returns the chosen icon number, or null on Cancel.</summary>
public partial class HotlineIconPickerWindow : Window
{
    private readonly HotlineIconPickerViewModel _viewModel = new();

    public HotlineIconPickerWindow()
    {
        InitializeComponent();
        DataContext = _viewModel;
        _viewModel.Picked = id => Close((ushort?)id);
        Opened += async (_, _) => await _viewModel.LoadAsync();
        Closed += (_, _) => _viewModel.Cancel();
    }

    private void OnCancelClick(object? sender, RoutedEventArgs e) => Close(null);
}
