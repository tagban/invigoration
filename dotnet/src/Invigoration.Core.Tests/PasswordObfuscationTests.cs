using Invigoration.Core.Config;
using Invigoration.Core.Crypto;

namespace Invigoration.Core.Tests;

public class PasswordObfuscatorTests
{
    [Theory]
    [InlineData("hunter2")]
    [InlineData("")]
    [InlineData("p@ss w0rd with spaces!")]
    public void Wrap_ThenUnwrap_RoundTrips(string password)
    {
        var wrapped = PasswordObfuscator.Wrap(password);

        Assert.Equal(password, PasswordObfuscator.Unwrap(wrapped));
    }

    [Fact]
    public void Wrap_NonEmptyPassword_IsBracketedAndNotPlaintext()
    {
        var wrapped = PasswordObfuscator.Wrap("hunter2");

        Assert.StartsWith("[", wrapped);
        Assert.EndsWith("]", wrapped);
        Assert.DoesNotContain("hunter2", wrapped);
    }

    [Fact]
    public void Unwrap_PlaintextWithoutBrackets_PassesThroughUnchanged()
    {
        // A password typed directly into bots.json by hand should work immediately.
        Assert.Equal("mynewpassword", PasswordObfuscator.Unwrap("mynewpassword"));
    }
}

public class ConfigStorePasswordPersistenceTests
{
    [Fact]
    public void SaveThenLoad_ObfuscatesOnDiskButRoundTripsPlaintext()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invig-test-{Guid.NewGuid():N}.json");
        try
        {
            var store = new ConfigStore(path);
            var bot = new BotConfig { DisplayName = "Test", Password = "hunter2" };
            store.Save([bot]);

            var rawJson = File.ReadAllText(path);
            Assert.DoesNotContain("hunter2", rawJson);

            var loaded = store.Load();
            Assert.Equal("hunter2", Assert.Single(loaded).Password);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Load_HandEditedPlaintextPassword_WorksAndGetsWrappedOnNextSave()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invig-test-{Guid.NewGuid():N}.json");
        try
        {
            File.WriteAllText(path, """[{"DisplayName":"Test","Password":"manuallyTyped"}]""");
            var store = new ConfigStore(path);

            var loaded = store.Load();
            Assert.Equal("manuallyTyped", Assert.Single(loaded).Password);

            store.Save(loaded);
            var rawJson = File.ReadAllText(path);
            Assert.DoesNotContain("manuallyTyped", rawJson);
            Assert.Equal("manuallyTyped", Assert.Single(store.Load()).Password);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Clone_PreservesPlaintextPassword()
    {
        var original = new BotConfig { Password = "hunter2" };

        var clone = BotConfig.Clone(original);

        Assert.Equal("hunter2", clone.Password);
    }
}
