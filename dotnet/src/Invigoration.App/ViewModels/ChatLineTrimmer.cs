using System.Collections.ObjectModel;
using System.Collections.Specialized;
using Avalonia.Threading;
using Invigoration.App.Models;

namespace Invigoration.App.ViewModels;

/// <summary>
/// Caps a chat log's ObservableCollection&lt;ChatLineViewModel&gt; at a bounded size, trimming the
/// oldest lines in batches once it grows past the cap. An unbounded chat log (join/leave lines
/// during sustained load testing routinely reaching into the thousands) meant the single
/// SelectableTextBlock rendering it (BotTabView/ChannelTabView's AppendLine, appending to one
/// flat Inlines collection with no virtualization) had to re-measure an ever-larger Inlines
/// collection on every new line — confirmed live as a real cause of the UI lagging worse the
/// longer a session ran, independent of every earlier backend-side fix (roster lookups, send
/// queue pileup). Trimming in batches, not down to the cap on every single overflow, means the
/// visual rebuild this forces (the onTrimmed callback) only fires once every TrimBatchSize lines
/// rather than on every one past the cap.
/// </summary>
public static class ChatLineTrimmer
{
    public const int MaxLines = 500;
    private const int TrimBatchSize = 100;

    /// <summary>
    /// Wire this once per chat-log collection (typically from the owning ViewModel's
    /// constructor). onTrimmed fires after the collection already holds its final trimmed
    /// contents, so the view can safely do a full (now bounded — at most MaxLines items — so
    /// genuinely cheap) Inlines rebuild rather than trying to reconcile the several RemoveAt
    /// notifications the trim itself raises.
    ///
    /// The actual trim is deferred to the next UI dispatch tick (Dispatcher.UIThread.Post)
    /// rather than run synchronously inside the CollectionChanged handler that detects the
    /// overflow — confirmed live as necessary, not just cautious: .NET invokes every subscriber
    /// of the same event in subscription order, and this handler is attached before the view's
    /// own (which appends the just-added line's Inlines in response to that same Add
    /// notification). Trimming synchronously here — including firing onTrimmed, which rebuilds
    /// Inlines from the collection's current contents — would run and complete *before* the
    /// view's own Add handler gets its turn on the very notification that triggered the trim, so
    /// the newest line would render once from the rebuild and then a second time when the view's
    /// normal handler finally runs. Deferring lets the original Add's full subscriber chain (the
    /// view's normal single-line append included) finish first.
    /// </summary>
    public static void Attach(ObservableCollection<ChatLineViewModel> lines, Action onTrimmed)
    {
        lines.CollectionChanged += (_, e) =>
        {
            if (e.Action != NotifyCollectionChangedAction.Add || lines.Count <= MaxLines + TrimBatchSize)
            {
                return;
            }

            Dispatcher.UIThread.Post(() =>
            {
                // Re-check: another Add already queued (and by now possibly already ran) its own
                // trim for the same overflow by the time this posted callback actually runs.
                if (lines.Count <= MaxLines + TrimBatchSize)
                {
                    return;
                }

                var excess = lines.Count - MaxLines;
                for (var i = 0; i < excess; i++)
                {
                    lines.RemoveAt(0);
                }

                onTrimmed();
            });
        };
    }
}
