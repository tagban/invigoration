using Invigoration.Core.Networking;

namespace Invigoration.Core.Chat;

/// <summary>Parses a raw SID_CHATEVENT frame. Port of bnetbot.cls's DispatchMessage.</summary>
public static class ChatEventParser
{
    public static ChatEvent Parse(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        var eventId = (ChatEventType)reader.ReadDword();
        var flags = reader.ReadDword();
        var ping = reader.ReadDword();
        reader.Skip(12); // IP, Account Number, Registration Authority: deprecated/unused fields
        var username = reader.ReadNTString();
        var text = reader.ReadNTString();
        return new ChatEvent(eventId, username, flags, (int)ping, text);
    }
}
