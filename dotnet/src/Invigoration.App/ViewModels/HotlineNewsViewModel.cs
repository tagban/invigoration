using System.Collections.ObjectModel;
using Avalonia;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.Core.Hotline;

namespace Invigoration.App.ViewModels;

/// <summary>A bundle or category row in the news tree.</summary>
public sealed class HotlineNewsCategoryRow(HotlineNewsCategory category)
{
    public HotlineNewsCategory Category { get; } = category;

    public string Name => Category.Name;

    public string Icon => Category.IsBundle ? "📦" : "📰";

    public string CountText => Category.IsBundle
        ? ""
        : $"{Category.ArticleCount} article{(Category.ArticleCount == 1 ? "" : "s")}";
}

/// <summary>One article in the list.</summary>
public sealed class HotlineNewsArticleRow(HotlineNewsArticle article)
{
    public HotlineNewsArticle Article { get; } = article;

    /// <summary>Replies are indented one step so a thread reads as a thread — the protocol gives a parent ID, not a depth, so this is one level, not a full tree.</summary>
    public Thickness Indent => new(Article.ParentId != 0 ? 20 : 0, 0, 0, 0);

    public string Title => string.IsNullOrWhiteSpace(Article.Title) ? "(no subject)" : Article.Title;

    public string Byline => Article.Posted is { } posted
        ? $"{Article.Poster} — {posted.LocalDateTime:g}"
        : Article.Poster;
}

/// <summary>
/// The News tab of one connected server: walk the category tree, read an article, and post one
/// (or reply). Posting is hidden entirely for an account without the privilege, rather than
/// offering a button that only ever fails.
/// </summary>
public sealed partial class HotlineNewsViewModel(HotlineTransactionClient client) : ObservableObject
{
    private readonly List<string> _path = [];

    public ObservableCollection<HotlineNewsCategoryRow> Categories { get; } = [];

    public ObservableCollection<HotlineNewsArticleRow> Articles { get; } = [];

    public string PathText => _path.Count == 0 ? "News" : "News / " + string.Join(" / ", _path);

    public bool CanGoUp => _path.Count > 0;

    /// <summary>At the root a post goes to the flat news document; inside a category it's a threaded article. Either way the account needs the privilege.</summary>
    public bool CanPost => client.CanPostNews;

    /// <summary>True when a post would go to the flat news rather than a threaded category — the composer says so, since the two land in different places.</summary>
    public bool PostsToFlatNews => _path.Count == 0;

    public string ComposeHeading => PostsToFlatNews ? "Post to the server's news" : $"Post to {_path[^1]}";

    /// <summary>
    /// The server's flat news — the original 1.x form, one text document rather than a tree. On
    /// most classic servers this is where the real news actually is, so it's shown first and the
    /// category tree (if any) below it.
    /// </summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasFlatNews))]
    [NotifyPropertyChangedFor(nameof(ShowEmptyHint))]
    public partial string FlatNews { get; set; } = "";

    public bool HasFlatNews => !string.IsNullOrWhiteSpace(FlatNews);

    /// <summary>Nothing has been loaded yet (or there's genuinely nothing here) — shows a hint instead of an unexplained blank pane. Never while a request is in flight.</summary>
    public bool ShowEmptyHint => !IsBusy && !HasFlatNews && Categories.Count == 0 && Articles.Count == 0;

    /// <summary>True once we're inside a category — that's when the article list, rather than the category list, is what to show.</summary>
    public bool IsInCategory => Articles.Count > 0 || (_path.Count > 0 && Categories.Count == 0);

    [ObservableProperty]
    public partial bool IsBusy { get; set; }

    [ObservableProperty]
    public partial string StatusMessage { get; set; } = "";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasOpenArticle))]
    public partial HotlineNewsArticleRow? SelectedArticle { get; set; }

    [ObservableProperty]
    public partial string ArticleBody { get; set; } = "";

    public bool HasOpenArticle => SelectedArticle is not null;

    // --- Composing ---

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsComposing))]
    public partial bool ComposeOpen { get; set; }

    public bool IsComposing => ComposeOpen;

    [ObservableProperty]
    public partial string ComposeTitle { get; set; } = "";

    [ObservableProperty]
    public partial string ComposeBody { get; set; } = "";

    /// <summary>Non-zero when the composer is replying to an article rather than starting a thread.</summary>
    private uint _replyingTo;

    partial void OnSelectedArticleChanged(HotlineNewsArticleRow? value)
    {
        ArticleBody = "";
        if (value is not null)
        {
            _ = LoadBodyAsync(value);
        }
    }

    private async Task LoadBodyAsync(HotlineNewsArticleRow row)
    {
        var body = await client.GetNewsArticleBodyAsync(_path, row.Article.Id, row.Article.Flavor).ConfigureAwait(true);

        // Guard against a slow fetch landing after the user moved on to another article.
        if (ReferenceEquals(SelectedArticle, row))
        {
            ArticleBody = body ?? "(couldn't read this article)";
        }
    }

    [RelayCommand]
    public async Task RefreshAsync()
    {
        if (IsBusy)
        {
            return;
        }

        IsBusy = true;
        StatusMessage = "";
        SelectedArticle = null;
        try
        {
            if (!client.CanReadNews)
            {
                StatusMessage = "This account isn't allowed to read news on this server.";
                return;
            }

            // Only at the root: inside a category the flat document isn't what's being looked at.
            if (_path.Count == 0)
            {
                FlatNews = await client.GetFlatNewsAsync().ConfigureAwait(true) ?? "";
            }

            var categories = await client.GetNewsCategoriesAsync(_path).ConfigureAwait(true);
            Categories.Clear();
            foreach (var category in categories)
            {
                Categories.Add(new HotlineNewsCategoryRow(category));
            }

            // A path with no categories under it is a category itself — ask for its articles. At
            // the root this never happens, which is why it's keyed on the path being non-empty.
            Articles.Clear();
            if (categories.Count == 0 && _path.Count > 0)
            {
                foreach (var article in await client.GetNewsArticlesAsync(_path).ConfigureAwait(true))
                {
                    Articles.Add(new HotlineNewsArticleRow(article));
                }
            }

            if (Categories.Count == 0 && Articles.Count == 0 && !HasFlatNews)
            {
                StatusMessage = _path.Count == 0 ? "This server has no news." : "Nothing posted here yet.";
            }
        }
        catch (Exception ex) when (ex is IOException or InvalidOperationException)
        {
            StatusMessage = $"Couldn't load news: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
            RaiseLocationChanged();
        }
    }

    [RelayCommand]
    private async Task OpenCategoryAsync(HotlineNewsCategoryRow? row)
    {
        if (row is null)
        {
            return;
        }

        _path.Add(row.Name);
        await RefreshAsync().ConfigureAwait(true);
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

    [RelayCommand]
    private void StartPost()
    {
        _replyingTo = 0;
        ComposeTitle = "";
        ComposeBody = "";
        ComposeOpen = true;
    }

    [RelayCommand]
    private void StartReply()
    {
        if (SelectedArticle is not { } selected)
        {
            return;
        }

        _replyingTo = selected.Article.Id;
        ComposeTitle = selected.Article.Title.StartsWith("Re:", StringComparison.OrdinalIgnoreCase)
            ? selected.Article.Title
            : $"Re: {selected.Article.Title}";
        ComposeBody = "";
        ComposeOpen = true;
    }

    [RelayCommand]
    private void CancelPost() => ComposeOpen = false;

    [RelayCommand]
    private async Task SubmitPostAsync()
    {
        if (string.IsNullOrWhiteSpace(ComposeTitle))
        {
            StatusMessage = "Give the post a subject first.";
            return;
        }

        var posted = PostsToFlatNews
            ? await client.PostFlatNewsAsync(FormatFlatPost(ComposeTitle.Trim(), ComposeBody)).ConfigureAwait(true)
            : await client.PostNewsArticleAsync(_path, ComposeTitle.Trim(), ComposeBody, _replyingTo).ConfigureAwait(true);
        if (!posted)
        {
            StatusMessage = "The server wouldn't accept the post.";
            return;
        }

        ComposeOpen = false;
        StatusMessage = "Posted.";
        await RefreshAsync().ConfigureAwait(true);
    }

    /// <summary>Puts a just-posted item on top of the flat news, the way the server itself orders it (newest first).</summary>
    public void PrependFlatNews(string post)
    {
        FlatNews = string.IsNullOrEmpty(FlatNews) ? post : post + "\r" + FlatNews;
    }

    /// <summary>
    /// The flat news is one plain document with no fields of its own, so a subject only exists if
    /// it's written into the text. Real clients put it on the first line; this matches that rather
    /// than silently dropping what the user typed in the subject box.
    /// </summary>
    private static string FormatFlatPost(string title, string body) =>
        string.IsNullOrWhiteSpace(title) ? body : $"{title}\r\r{body}";

    private void RaiseLocationChanged()
    {
        OnPropertyChanged(nameof(PathText));
        OnPropertyChanged(nameof(CanGoUp));
        OnPropertyChanged(nameof(CanPost));
        OnPropertyChanged(nameof(IsInCategory));
        OnPropertyChanged(nameof(ShowEmptyHint));
        OnPropertyChanged(nameof(PostsToFlatNews));
        OnPropertyChanged(nameof(ComposeHeading));
    }
}
