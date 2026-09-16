using System.Collections.ObjectModel;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.Core.Hotline;

namespace Invigoration.App.ViewModels;

/// <summary>
/// One row in the file list — a folder to open or a file to download.
///
/// The row carries its own Open/Download commands, rather than the list binding to the files view
/// model with a CommandParameter, because a context menu lives in its own popup tree: a binding
/// that walks up to the parent ListBox resolves to nothing from in there. Hanging the commands off
/// the row means the menu binds to its own DataContext and needs no such walk.
/// </summary>
public sealed partial class HotlineFileRowViewModel(HotlineFileEntry entry, HotlineFilesViewModel owner) : ObservableObject
{
    public HotlineFileEntry Entry { get; } = entry;

    public string Name => Entry.Name;

    public bool IsFolder => Entry.IsFolder;

    /// <summary>What the menu offers for this row — only a folder can be browsed into.</summary>
    public string DownloadLabel => IsFolder ? "Download Folder..." : "Download...";

    /// <summary>Double-clicking, or picking Open: a folder browses in, anything else downloads.</summary>
    [RelayCommand]
    public Task ActivateAsync() => owner.OpenAsync(this);

    [RelayCommand]
    public Task DownloadAsync() => owner.DownloadAsync(this);

    /// <summary>Whether this account may delete this particular row — the protocol tracks files and folders as separate privileges.</summary>
    public bool CanDelete => owner.CanDelete(IsFolder);

    public bool CanRename => owner.CanRename(IsFolder);

    [RelayCommand]
    private void Delete() => owner.AskToDelete(this);

    [RelayCommand]
    private void Rename() => owner.AskToRename(this);

    public string Icon => Entry.IsFolder ? "📁" : "📄";

    /// <summary>A folder's "size" is how many items it holds, not bytes — showing it as bytes would be a lie.</summary>
    public string SizeText => Entry.IsFolder
        ? $"{Entry.Size} item{(Entry.Size == 1 ? "" : "s")}"
        : FormatBytes(Entry.Size);

    /// <summary>Progress of this row's own download, 0 when nothing's running.</summary>
    [ObservableProperty]
    public partial double DownloadProgress { get; set; }

    [ObservableProperty]
    public partial string Status { get; set; } = "";

    internal static string FormatBytes(long bytes) => bytes switch
    {
        < 1024 => $"{bytes} B",
        < 1024 * 1024 => $"{bytes / 1024.0:0.#} KB",
        < 1024L * 1024 * 1024 => $"{bytes / (1024.0 * 1024):0.#} MB",
        _ => $"{bytes / (1024.0 * 1024 * 1024):0.##} GB",
    };
}

/// <summary>
/// The Files tab of one connected server: browse the server's folders, download a file to disk,
/// upload one from it. Every request goes through the session's own live connection; the bytes
/// themselves move on their own connection (see HotlineFileTransfer), so a transfer in progress
/// doesn't block chat.
/// </summary>
public sealed partial class HotlineFilesViewModel(HotlineTransactionClient client, string host, int port) : ObservableObject
{
    private readonly List<string> _path = [];

    public ObservableCollection<HotlineFileRowViewModel> Entries { get; } = [];

    /// <summary>Where we are, for the header — "Files" at the root, otherwise the path joined.</summary>
    public string PathText => _path.Count == 0 ? "Files" : "Files / " + string.Join(" / ", _path);

    public bool CanGoUp => _path.Count > 0;

    /// <summary>Always offered — see the remarks on CanDelete: upload is set per folder, and a drop box is exactly the case where the account-wide bit says no and the folder says yes.</summary>
    public bool CanUpload => true;

    /// <summary>Nothing loaded yet (or the folder is genuinely empty) — a hint beats an unexplained blank pane. Never while a request is in flight.</summary>
    public bool ShowEmptyHint => !IsBusy && Entries.Count == 0;

    [ObservableProperty]
    public partial bool IsBusy { get; set; }

    [ObservableProperty]
    public partial string StatusMessage { get; set; } = "";

    [ObservableProperty]
    public partial HotlineFileRowViewModel? SelectedEntry { get; set; }

    /// <summary>What the inline prompt strip at the bottom of the Files tab is currently asking.</summary>
    public enum PromptKind
    {
        None,
        NewFolder,
        Rename,
        ConfirmDelete,
    }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsPrompting))]
    [NotifyPropertyChangedFor(nameof(NeedsPromptText))]
    public partial PromptKind Prompt { get; set; }

    public bool IsPrompting => Prompt != PromptKind.None;

    /// <summary>A delete only needs a yes; the other two need a name typed.</summary>
    public bool NeedsPromptText => Prompt is PromptKind.NewFolder or PromptKind.Rename;

    [ObservableProperty]
    public partial string PromptHeading { get; set; } = "";

    [ObservableProperty]
    public partial string PromptText { get; set; } = "";

    /// <summary>The verb on the confirm button — "Create", "Rename", "Delete" — so it says what will happen rather than "OK".</summary>
    [ObservableProperty]
    public partial string PromptAction { get; set; } = "";

    private HotlineFileRowViewModel? _promptTarget;

    // Hotline sets Upload, Download, Delete File, Create Folder and Delete Folder PER FOLDER as
    // well as per account — a drop box grants uploads to people who don't have the privilege
    // generally, and a locked folder refuses deletes from people who do. The protocol offers no
    // way to ask what a given folder allows: a client finds out by trying.
    //
    // So the account bitmap is a hint, not a gate. Every action is offered and the server decides;
    // a refusal comes back as a plain message, with a note when the account lacks the privilege
    // generally (see DescribeRefusal), since that's the likeliest explanation. Hiding these
    // instead would make a drop box impossible to use from this client.
    public bool CanDelete(bool isFolder) => true;

    public bool CanRename(bool isFolder) => true;

    public bool CanCreateFolder => true;

    /// <summary>Why the server probably said no — the account-wide privilege when it's missing, since a folder can only ever have been the other reason.</summary>
    private string DescribeRefusal(string action, bool hasAccountPrivilege) => hasAccountPrivilege
        ? $"The server wouldn't {action} that — this folder may not allow it."
        : $"The server wouldn't {action} that. This account doesn't have that privilege generally, though some folders grant it individually.";

    [RelayCommand]
    private void StartNewFolder()
    {
        _promptTarget = null;
        Prompt = PromptKind.NewFolder;
        PromptHeading = $"New folder in {PathText}";
        PromptAction = "Create";
        PromptText = "";
    }

    public void AskToRename(HotlineFileRowViewModel row)
    {
        _promptTarget = row;
        Prompt = PromptKind.Rename;
        PromptHeading = $"Rename \"{row.Name.Trim()}\"";
        PromptAction = "Rename";
        PromptText = row.Name.Trim();
    }

    /// <summary>
    /// Deleting asks first, and names what will go. It happens on the server, to everyone, with no
    /// undo in the protocol — a misclick on a folder full of other people's files is not
    /// recoverable from here.
    /// </summary>
    public void AskToDelete(HotlineFileRowViewModel row)
    {
        _promptTarget = row;
        Prompt = PromptKind.ConfirmDelete;
        PromptHeading = row.IsFolder
            ? $"Delete the folder \"{row.Name.Trim()}\" and everything in it? This can't be undone."
            : $"Delete \"{row.Name.Trim()}\"? This can't be undone.";
        PromptAction = "Delete";
        PromptText = "";
    }

    [RelayCommand]
    private void CancelPrompt()
    {
        Prompt = PromptKind.None;
        _promptTarget = null;
    }

    [RelayCommand]
    private async Task SubmitPromptAsync()
    {
        var kind = Prompt;
        var target = _promptTarget;
        var typed = PromptText.Trim();
        Prompt = PromptKind.None;
        _promptTarget = null;

        try
        {
            switch (kind)
            {
                case PromptKind.NewFolder when typed.Length > 0:
                    StatusMessage = await client.CreateFolderAsync(_path, typed).ConfigureAwait(true)
                        ? $"Created {typed}."
                        : DescribeRefusal("create", client.CanCreateFolders);
                    break;

                case PromptKind.Rename when target is not null && typed.Length > 0:
                    StatusMessage = await client.RenameAsync(_path, target.Name, typed).ConfigureAwait(true)
                        ? $"Renamed to {typed}."
                        : DescribeRefusal("rename", target.IsFolder ? client.CanRenameFolders : client.CanRenameFiles);
                    break;

                case PromptKind.ConfirmDelete when target is not null:
                    StatusMessage = await client.DeleteFileAsync(_path, target.Name).ConfigureAwait(true)
                        ? $"Deleted {target.Name.Trim()}."
                        : DescribeRefusal("delete", target.IsFolder ? client.CanDeleteFolders : client.CanDeleteFiles);
                    break;

                default:
                    return;
            }

            await RefreshAsync().ConfigureAwait(true);
        }
        catch (Exception ex) when (ex is IOException or System.Net.Sockets.SocketException)
        {
            StatusMessage = $"Failed: {ex.Message}";
        }
    }

    /// <summary>Set by the view: asks the user where to save a download, or which file to upload. Null result means they cancelled.</summary>
    public Func<string, Task<string?>>? AskWhereToSave { get; set; }

    public Func<Task<string?>>? AskWhatToUpload { get; set; }

    [RelayCommand]
    public async Task RefreshAsync()
    {
        if (IsBusy)
        {
            return;
        }

        IsBusy = true;
        StatusMessage = "";
        try
        {
            var entries = await client.GetFileListAsync(_path).ConfigureAwait(true);
            Entries.Clear();

            // Folders first, then files, each alphabetically — the order every file browser uses,
            // and not something the server guarantees.
            foreach (var entry in entries.OrderByDescending(e => e.IsFolder).ThenBy(e => e.Name, StringComparer.OrdinalIgnoreCase))
            {
                Entries.Add(new HotlineFileRowViewModel(entry, this));
            }

            if (Entries.Count == 0)
            {
                StatusMessage = "Nothing here.";
            }
        }
        catch (Exception ex) when (ex is IOException or InvalidOperationException)
        {
            StatusMessage = $"Couldn't list this folder: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
            OnPropertyChanged(nameof(PathText));
            OnPropertyChanged(nameof(CanGoUp));
            OnPropertyChanged(nameof(ShowEmptyHint));
        }
    }

    [RelayCommand]
    public async Task OpenAsync(HotlineFileRowViewModel? row)
    {
        if (row is null)
        {
            return;
        }

        if (row.IsFolder)
        {
            _path.Add(row.Name);
            await RefreshAsync().ConfigureAwait(true);
        }
        else
        {
            await DownloadAsync(row).ConfigureAwait(true);
        }
    }

    [RelayCommand]
    private async Task GoUpAsync()
    {
        if (_path.Count == 0)
        {
            return;
        }

        _path.RemoveAt(_path.Count - 1);
        await RefreshAsync().ConfigureAwait(true);
    }

    /// <summary>Set by the view: asks which folder to download into. Null means cancelled.</summary>
    public Func<string, Task<string?>>? AskWhereToSaveFolder { get; set; }

    [RelayCommand]
    public async Task DownloadAsync(HotlineFileRowViewModel? row)
    {
        row ??= SelectedEntry;
        if (row is null)
        {
            return;
        }

        if (row.IsFolder)
        {
            await DownloadFolderAsync(row).ConfigureAwait(true);
            return;
        }

        if (AskWhereToSave is null)
        {
            return;
        }

        var destination = await AskWhereToSave(row.Name).ConfigureAwait(true);
        if (destination is null)
        {
            return;
        }

        row.Status = "Starting...";
        try
        {
            var ticket = await client.RequestDownloadAsync(_path, row.Name).ConfigureAwait(true);
            if (ticket is null)
            {
                row.Status = "The server wouldn't send this file.";
                return;
            }

            if (ticket.WaitingCount > 0)
            {
                row.Status = $"Queued behind {ticket.WaitingCount}...";
            }

            var progress = new Progress<HotlineTransferProgress>(p => Dispatcher.UIThread.Post(() =>
            {
                row.DownloadProgress = p.Fraction * 100 ?? 0;
                row.Status = p.Total > 0
                    ? $"{HotlineFileRowViewModel.FormatBytes(p.Transferred)} of {HotlineFileRowViewModel.FormatBytes(p.Total)}"
                    : HotlineFileRowViewModel.FormatBytes(p.Transferred);
            }));

            // Written to a temporary name first so a failed transfer can't leave a half-file
            // sitting there looking finished.
            var partial = destination + ".part";
            await using (var file = File.Create(partial))
            {
                await HotlineFileTransfer.DownloadAsync(host, port, ticket.ReferenceNumber, file, progress).ConfigureAwait(true);
            }

            File.Move(partial, destination, overwrite: true);
            row.DownloadProgress = 100;
            row.Status = "Saved.";
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or System.Net.Sockets.SocketException)
        {
            row.DownloadProgress = 0;
            row.Status = $"Failed: {ex.Message}";
            TryDeletePartial(destination + ".part");
        }
    }

    /// <summary>
    /// Downloads a whole folder, recreating its structure under a directory the user picks. The
    /// server walks its contents item by item (see HotlineFolderTransfer) — the item count from
    /// the request is how both sides know when it's done.
    /// </summary>
    private async Task DownloadFolderAsync(HotlineFileRowViewModel row)
    {
        if (AskWhereToSaveFolder is null)
        {
            return;
        }

        var destination = await AskWhereToSaveFolder(row.Name).ConfigureAwait(true);
        if (destination is null)
        {
            return;
        }

        row.Status = "Starting...";
        try
        {
            var ticket = await client.RequestFolderDownloadAsync(_path, row.Name).ConfigureAwait(true);
            if (ticket is null)
            {
                row.Status = "The server wouldn't send this folder.";
                return;
            }

            if (ticket.ItemCount == 0)
            {
                row.Status = "The server says this folder is empty.";
                return;
            }

            var done = 0;
            var progress = new Progress<HotlineFolderItem>(item => Dispatcher.UIThread.Post(() =>
            {
                if (!item.IsFolder)
                {
                    done++;
                }

                row.DownloadProgress = ticket.ItemCount > 0 ? Math.Min(100.0 * done / ticket.ItemCount, 100) : 0;
                row.Status = $"{item.Name} ({done} of {ticket.ItemCount})";
            }));

            // Into a subfolder named after the folder itself, so picking a destination twice
            // doesn't scatter two servers' folders through the same directory.
            var into = Path.Combine(destination, HotlineFolderTransfer.SafePath([row.Name]).FirstOrDefault() ?? "download");
            Directory.CreateDirectory(into);

            var written = await HotlineFolderTransfer.DownloadAsync(
                host, port, ticket.ReferenceNumber, ticket.ItemCount, into, progress).ConfigureAwait(true);

            row.DownloadProgress = 100;
            row.Status = $"Saved {written} file{(written == 1 ? "" : "s")}.";
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or System.Net.Sockets.SocketException)
        {
            row.DownloadProgress = 0;
            row.Status = $"Failed: {ex.Message}";
        }
    }

    /// <summary>Set by the view: asks which folder to upload. Null means cancelled.</summary>
    public Func<Task<string?>>? AskWhatFolderToUpload { get; set; }

    [RelayCommand]
    private async Task UploadFolderAsync()
    {
        if (AskWhatFolderToUpload is null)
        {
            return;
        }

        var source = await AskWhatFolderToUpload().ConfigureAwait(true);
        if (source is null || !Directory.Exists(source))
        {
            return;
        }

        var name = new DirectoryInfo(source).Name;
        var items = HotlineFolderTransfer.Enumerate(source);
        var totalBytes = items.Where(i => !i.IsFolder)
            .Sum(i => HotlineFileTransfer.FlattenedSize(i.Path[^1], i.Size));

        StatusMessage = $"Uploading {name} ({items.Count} items)...";
        try
        {
            var ticket = await client.RequestFolderUploadAsync(_path, name, items.Count, totalBytes).ConfigureAwait(true);
            if (ticket is null)
            {
                StatusMessage = DescribeRefusal("accept a folder upload of", client.CanUploadFiles);
                return;
            }

            var done = 0;
            var progress = new Progress<HotlineFolderItem>(item => Dispatcher.UIThread.Post(() =>
            {
                if (!item.IsFolder)
                {
                    done++;
                }

                StatusMessage = $"Uploading {name}: {item.Name} ({done} of {items.Count})";
            }));

            var sent = await HotlineFolderTransfer.UploadAsync(host, port, ticket.ReferenceNumber, source, progress)
                .ConfigureAwait(true);

            StatusMessage = $"Uploaded {name} — {sent} file{(sent == 1 ? "" : "s")}.";
            await RefreshAsync().ConfigureAwait(true);
        }
        catch (Exception ex) when (ex is IOException or System.Net.Sockets.SocketException)
        {
            StatusMessage = $"Folder upload failed: {ex.Message}";
        }
    }

    [RelayCommand]
    private async Task UploadAsync()
    {
        if (AskWhatToUpload is null)
        {
            return;
        }

        var source = await AskWhatToUpload().ConfigureAwait(true);
        if (source is null || !File.Exists(source))
        {
            return;
        }

        var name = Path.GetFileName(source);
        var length = new FileInfo(source).Length;
        StatusMessage = $"Uploading {name}...";

        try
        {
            var ticket = await client.RequestUploadAsync(_path, name, length).ConfigureAwait(true);
            if (ticket is null)
            {
                StatusMessage = DescribeRefusal("accept an upload of", client.CanUploadFiles);
                return;
            }

            var progress = new Progress<HotlineTransferProgress>(p => Dispatcher.UIThread.Post(() =>
                StatusMessage = $"Uploading {name}: {HotlineFileRowViewModel.FormatBytes(p.Transferred)} of {HotlineFileRowViewModel.FormatBytes(p.Total)}"));

            await using (var file = File.OpenRead(source))
            {
                await HotlineFileTransfer.UploadAsync(host, port, ticket.ReferenceNumber, name, file, length, progress).ConfigureAwait(true);
            }

            StatusMessage = $"Uploaded {name}.";
            await RefreshAsync().ConfigureAwait(true);
        }
        catch (Exception ex) when (ex is IOException or System.Net.Sockets.SocketException)
        {
            StatusMessage = $"Upload failed: {ex.Message}";
        }
    }

    private static void TryDeletePartial(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
            // Leaving a .part behind is not worth reporting on top of the failure that caused it.
        }
    }
}
