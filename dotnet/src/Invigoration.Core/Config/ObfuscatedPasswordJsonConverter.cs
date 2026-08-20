using System.Text.Json;
using System.Text.Json.Serialization;
using Invigoration.Core.Crypto;

namespace Invigoration.Core.Config;

/// <summary>
/// Applied to BotConfig.Password so it's obfuscated at rest wherever this
/// config gets JSON-serialized (ConfigStore.Save, and BotConfig.Clone's
/// serialize/deserialize round-trip) without the in-memory Password value
/// ever needing to be anything but plaintext.
/// </summary>
public sealed class ObfuscatedPasswordJsonConverter : JsonConverter<string>
{
    public override string Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options) =>
        PasswordObfuscator.Unwrap(reader.GetString() ?? "");

    public override void Write(Utf8JsonWriter writer, string value, JsonSerializerOptions options) =>
        writer.WriteStringValue(PasswordObfuscator.Wrap(value));
}
