using System.Diagnostics;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Threading;
using Invigoration.App.Models;

namespace Invigoration.App.Views;

/// <summary>
/// Plays a <see cref="SpriteAnimation"/>, pixel-exact. Every instance runs off one shared clock that
/// only ticks while at least one is on screen, and only redraws when its frame actually changes — the
/// character dock is a virtualizing list, so this stays cheap however many people are in the channel.
/// <see cref="Phase"/> offsets where in the loop an instance starts, so a crowd of the same avatar
/// doesn't move in lockstep.
/// </summary>
public sealed class SpriteAnimationView : Control
{
    public static readonly StyledProperty<SpriteAnimation?> AnimationProperty =
        AvaloniaProperty.Register<SpriteAnimationView, SpriteAnimation?>(nameof(Animation));

    public static readonly StyledProperty<int> PhaseProperty =
        AvaloniaProperty.Register<SpriteAnimationView, int>(nameof(Phase));

    private static readonly Stopwatch Clock = Stopwatch.StartNew();
    private static readonly HashSet<SpriteAnimationView> Live = [];
    private static readonly DispatcherTimer Ticker = new(TimeSpan.FromMilliseconds(40), DispatcherPriority.Render, (_, _) => Tick());

    private int _shownFrame = -1;

    static SpriteAnimationView()
    {
        AffectsMeasure<SpriteAnimationView>(AnimationProperty);
        AffectsRender<SpriteAnimationView>(AnimationProperty, PhaseProperty);
    }

    public SpriteAnimation? Animation
    {
        get => GetValue(AnimationProperty);
        set => SetValue(AnimationProperty, value);
    }

    public int Phase
    {
        get => GetValue(PhaseProperty);
        set => SetValue(PhaseProperty, value);
    }

    protected override Size MeasureOverride(Size availableSize) =>
        Animation is { } a ? new Size(a.FrameWidth, a.FrameHeight) : default;

    protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnAttachedToVisualTree(e);
        Live.Add(this);
        if (!Ticker.IsEnabled)
        {
            Ticker.Start();
        }
    }

    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnDetachedFromVisualTree(e);
        Live.Remove(this);
        if (Live.Count == 0)
        {
            Ticker.Stop();
        }
    }

    public override void Render(DrawingContext context)
    {
        if (Animation is not { } a)
        {
            return;
        }

        _shownFrame = a.FrameAt(Clock.ElapsedMilliseconds + Phase);
        var source = new Rect(_shownFrame * a.FrameWidth, 0, a.FrameWidth, a.FrameHeight);
        using (context.PushRenderOptions(new RenderOptions { BitmapInterpolationMode = Avalonia.Media.Imaging.BitmapInterpolationMode.None }))
        {
            context.DrawImage(a.Sheet, source, new Rect(Bounds.Size));
        }
    }

    private static void Tick()
    {
        var now = Clock.ElapsedMilliseconds;
        foreach (var view in Live)
        {
            if (view.Animation is { } a && view.IsEffectivelyVisible && a.FrameAt(now + view.Phase) != view._shownFrame)
            {
                view.InvalidateVisual();
            }
        }
    }
}
