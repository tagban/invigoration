using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.Core.Chat;

namespace Invigoration.App.Models;

public partial class ChannelUserViewModel(string username) : ObservableObject
{
    public string Username { get; } = username;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(StatusIconImage))]
    public partial uint Flags { get; set; }

    [ObservableProperty]
    public partial int Ping { get; set; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ProductIconImage))]
    public partial string StatString { get; set; } = "";

    public Bitmap? ProductIconImage => GameIconLoader.Get(ChatIcon.GetProductIconKey(StatString));

    public Bitmap? StatusIconImage => GameIconLoader.Get(ChatIcon.GetStatusIconKey(Flags));
}
