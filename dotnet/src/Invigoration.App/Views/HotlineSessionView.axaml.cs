using System.Collections.Specialized;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Avalonia.VisualTree;
using Avalonia.Threading;
using Invigoration.App.ViewModels;

namespace Invigoration.App.Views;

public partial class HotlineSessionView : UserControl
{
    private HotlineSessionViewModel? _attachedViewModel;

    public HotlineSessionView()
    {
        InitializeComponent();

        var inputBox = this.FindControl<TextBox>("InputBox");
        if (inputBox is not null)
        {
            inputBox.KeyDown += (_, e) =>
            {
                if (e.Key == Key.Enter && DataContext is HotlineSessionViewModel vm)
                {
                    vm.SendCommand.Execute(null);
                    e.Handled = true;
                }
            };
        }

        DataContextChanged += (_, _) =>
        {
            AttachAutoScroll();
            AttachFilePickers();
        };
    }

    /// <summary>"Choose..." in the header — the picker window hands back an icon number, which only fills the Icon box; Apply is still what sends it.</summary>
    private async void OnChooseIconClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not HotlineSessionViewModel vm || TopLevel.GetTopLevel(this) is not Window owner)
        {
            return;
        }

        if (await new HotlineIconPickerWindow().ShowDialog<ushort?>(owner) is { } iconId)
        {
            vm.EditIconId = iconId;
        }
    }

    /// <summary>
    /// Downloads and uploads need the platform file pickers, which need a TopLevel — a view
    /// concern, so the Files view model asks through these callbacks rather than reaching for one
    /// itself. Both return null when the user cancels, which the view model treats as "don't".
    /// </summary>
    private void AttachFilePickers()
    {
        if (DataContext is not HotlineSessionViewModel vm)
        {
            return;
        }

        vm.Files.AskWhereToSave = async suggestedName =>
        {
            var file = await TopLevel.GetTopLevel(this)!.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Save file from server",
                SuggestedFileName = suggestedName,
            });

            return file?.TryGetLocalPath();
        };

        vm.Files.AskWhereToSaveFolder = async folderName =>
        {
            var folders = await TopLevel.GetTopLevel(this)!.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
            {
                Title = $"Where should \"{folderName}\" be saved?",
                AllowMultiple = false,
            });

            return folders.Count > 0 ? folders[0].TryGetLocalPath() : null;
        };

        vm.Files.AskWhatFolderToUpload = async () =>
        {
            var folders = await TopLevel.GetTopLevel(this)!.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
            {
                Title = "Choose a folder to upload",
                AllowMultiple = false,
            });

            return folders.Count > 0 ? folders[0].TryGetLocalPath() : null;
        };

        vm.Files.AskWhatToUpload = async () =>
        {
            var files = await TopLevel.GetTopLevel(this)!.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "Choose a file to upload",
                AllowMultiple = false,
            });

            return files.Count > 0 ? files[0].TryGetLocalPath() : null;
        };
    }

    /// <summary>
    /// Double-clicking a file row activates it — a folder opens, anything else downloads (the row's
    /// own ActivateCommand decides which). The event is handled here rather than bound in XAML
    /// because DoubleTapped fires on whatever element inside the row was under the pointer (the
    /// name, the size, the icon), so the row has to be found by walking up to the nearest
    /// ListBoxItem rather than read off the sender.
    /// </summary>
    private void OnFilesDoubleTapped(object? sender, TappedEventArgs e)
    {
        if ((e.Source as Visual)?.FindAncestorOfType<ListBoxItem>(includeSelf: true) is not { DataContext: HotlineFileRowViewModel row })
        {
            return;
        }

        row.ActivateCommand.Execute(null);
        e.Handled = true;
    }

    /// <summary>
    /// Same fix as ChannelTabView.axaml.cs's own autoscroll — a plain ListBox never scrolls itself
    /// to a newly-added item, so the chat log otherwise stays wherever it last was as new lines
    /// keep arriving above the visible area. Fixed per direct user report ("AutoScroll is also not
    /// working on Hotline chats"). ScrollIntoView (not FindControl-ing the ListBox's own internal
    /// ScrollViewer and calling ScrollToEnd) since that's the documented Avalonia ListBox API for
    /// this and doesn't depend on reaching into the control's template.
    /// </summary>
    private void AttachAutoScroll()
    {
        if (_attachedViewModel is not null)
        {
            _attachedViewModel.Messages.CollectionChanged -= OnMessagesChanged;
        }

        _attachedViewModel = DataContext as HotlineSessionViewModel;
        if (_attachedViewModel is null)
        {
            return;
        }

        _attachedViewModel.Messages.CollectionChanged += OnMessagesChanged;
        ScrollToLastMessage();
    }

    private void OnMessagesChanged(object? sender, NotifyCollectionChangedEventArgs e) => ScrollToLastMessage();

    private void ScrollToLastMessage()
    {
        if (_attachedViewModel is not { Messages.Count: > 0 } vm)
        {
            return;
        }

        var listBox = this.FindControl<ListBox>("MessagesList");
        if (listBox is null)
        {
            return;
        }

        // Posted, not called inline — the ListBox needs a layout pass to realize the just-added
        // item's container before ScrollIntoView can actually find it.
        Dispatcher.UIThread.Post(() => listBox.ScrollIntoView(vm.Messages[^1]), DispatcherPriority.Background);
    }

    /// <summary>A ListBox only supports item selection, not text selection — this is the direct answer to "I can't copy from the chat log" rather than relying on click-drag text selection alone.</summary>
    private async void OnCopyLogClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not HotlineSessionViewModel vm)
        {
            return;
        }

        var clipboard = TopLevel.GetTopLevel(this)?.Clipboard;
        if (clipboard is not null)
        {
            await clipboard.SetTextAsync(string.Join(Environment.NewLine, vm.Messages.Select(m => m.FullText)));
        }
    }
}
