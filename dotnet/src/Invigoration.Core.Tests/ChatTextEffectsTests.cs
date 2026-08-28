using Invigoration.Core.Chat;

namespace Invigoration.Core.Tests;

public class ChatTextEffectsTests
{
    [Fact]
    public void Apply_NoTogglesOrText_ReturnsUnchanged()
    {
        Assert.Equal("hello world", ChatTextEffects.Apply("hello world", fuddMode: false, canadaMode: false, prependText: "", postpendText: ""));
    }

    [Fact]
    public void Apply_FuddMode_ReplacesRWithW()
    {
        Assert.Equal("hello wowld", ChatTextEffects.Apply("hello rorld", fuddMode: true, canadaMode: false, prependText: "", postpendText: ""));
    }

    [Fact]
    public void Apply_FuddMode_ReplacesUppercaseRToo()
    {
        Assert.Equal("Wed Wovew", ChatTextEffects.Apply("Red Rover", fuddMode: true, canadaMode: false, prependText: "", postpendText: ""));
    }

    [Fact]
    public void Apply_CanadaMode_AppendsEh()
    {
        Assert.Equal("nice, eh?", ChatTextEffects.Apply("nice", fuddMode: false, canadaMode: true, prependText: "", postpendText: ""));
    }

    [Fact]
    public void Apply_PrependAndPostpend_WrapText()
    {
        Assert.Equal(">> hi << bye", ChatTextEffects.Apply("hi", fuddMode: false, canadaMode: false, prependText: ">>", postpendText: "<< bye"));
    }

    [Fact]
    public void Apply_AllTogether_AppliesFuddCanadaThenWrapsWithPrependPostpend()
    {
        // Fudd/Canada transform the core text first, then prepend/postpend wrap the already-transformed
        // text — a signature-style postpend must never get its own R's turned into W's.
        var result = ChatTextEffects.Apply("run", fuddMode: true, canadaMode: true, prependText: "Sig:", postpendText: "-Rover");
        Assert.Equal("Sig: wun, eh? -Rover", result);
    }
}
