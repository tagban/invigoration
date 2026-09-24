using Invigoration.Sc2.Front;
using Invigoration.Sc2.Protobuf;
using Invigoration.Sc2.Wire;
using Attribute = Invigoration.Sc2.Front.Attribute;

namespace Invigoration.Sc2.Tests;

public class BgsFriendsPresenceTests
{
    private const uint Bn = BgsPresenceFields.BattleNetProgram;
    private const ulong AccountHigh = 0x0100000000000000;
    private const ulong GameAccountHigh = 0x0200000000000000;

    private static EntityId Account(ulong low) => new() { High = AccountHigh, Low = low };

    private static EntityId GameAccount(ulong low) => new() { High = GameAccountHigh | 0x5332, Low = low };

    private static BgsPresenceFieldOperation Set(uint group, uint field, Variant value, ulong? uniqueId = null) =>
        new() { Key = new BgsPresenceFieldKey { Program = Bn, Group = group, Field = field, UniqueId = uniqueId }, Value = value };

    private static BgsPresenceFieldOperation Clear(uint group, uint field, ulong? uniqueId = null) =>
        new() { Key = new BgsPresenceFieldKey { Program = Bn, Group = group, Field = field, UniqueId = uniqueId }, Value = new Variant(), Clear = true };

    [Theory]
    [InlineData(FrontServices.Friends, 0xA3DDB1BDu)]
    [InlineData(FrontServices.FriendsListener, 0x6F259A13u)]
    [InlineData(FrontServices.Presence, 0xFA0796FFu)]
    [InlineData(FrontServices.PresenceListener, 0x890AB85Fu)]
    [InlineData(FrontServices.ChannelListener, 0xBF8C8094u)]
    [InlineData(FrontServices.FriendsListenerV1, 0xA6717548u)]
    [InlineData(FrontServices.PresenceListenerV1, 0xE8836F50u)]
    [InlineData(FrontServices.ChannelListenerV1, 0xDA660990u)]
    public void ServiceHashes_MatchTrinityCoreGeneratedConstants(string name, uint expected)
    {
        Assert.Equal(expected, ServiceHash.Compute(name));
    }

    [Theory]
    [InlineData("S2", 0x5332u)]
    [InlineData("BN", 0x424Eu)]
    [InlineData("BSAp", 0x42534170u)]
    [InlineData("Fen", 0x46656Eu)]
    public void FourCc_DecodeInvertsEncode(string text, uint value)
    {
        Assert.Equal(value, FourCc.Encode(text));
        Assert.Equal(text, FourCc.Decode(value));
    }

    [Fact]
    public void FriendsSubscribeRequest_EncodesObjectIdAsField2()
    {
        var bytes = new BgsFriendsSubscribeRequest { ObjectId = 5 }.Encode();
        Assert.Equal(Convert.FromHexString("1005"), bytes);
        Assert.Equal(5ul, BgsFriendsSubscribeRequest.Decode(bytes).ObjectId);
    }

    [Fact]
    public void Friend_LegacyLayout_ReadsNamesFromFields6And7()
    {
        var friend = new BgsFriendMessage
        {
            AccountId = Account(42),
            Roles = [1, 2],
            Privileges = 7,
            FullName = "Jim Raynor",
            BattleTag = "Raynor#1234",
            Attributes = [new Attribute { Name = "note", Value = new Variant { StringValue = "hi" } }],
        };

        var decoded = BgsFriendMessage.Decode(friend.Encode());

        Assert.Equal(42ul, decoded.AccountId.Low);
        Assert.Equal("Jim Raynor", decoded.FullName);
        Assert.Equal("Raynor#1234", decoded.BattleTag);
        Assert.Null(decoded.CreationTime);
        Assert.Equal([1u, 2u], decoded.Roles);
        Assert.Equal(7ul, decoded.Privileges);
        Assert.Equal("hi", decoded.Attributes.Single().Value.StringValue);
    }

    [Fact]
    public void Friend_CurrentLayout_ReadsField6AsCreationTime()
    {
        var decoded = BgsFriendMessage.Decode(new BgsFriendMessage { AccountId = Account(1), CreationTime = 1_700_000_000 }.Encode());

        Assert.Equal(1_700_000_000ul, decoded.CreationTime);
        Assert.Null(decoded.FullName);
        Assert.Null(decoded.BattleTag);
    }

    [Fact]
    public void Friend_AcceptsUnpackedRoles()
    {
        var w = new ProtoWriter();
        w.WriteBytesField(1, Account(9).Encode());
        w.WriteUInt32(3, 1);
        w.WriteUInt32(3, 3);

        Assert.Equal([1u, 3u], BgsFriendMessage.Decode(w.ToArray()).Roles);
    }

    [Fact]
    public void SubscribeResponse_RoundTripsFriendsAndInvitations()
    {
        var response = new BgsFriendsSubscribeResponse
        {
            MaxFriends = 200,
            Friends = [new BgsFriendMessage { AccountId = Account(1) }, new BgsFriendMessage { AccountId = Account(2), BattleTag = "Kerrigan#1" }],
            ReceivedInvitations =
            [
                new BgsReceivedInvitation
                {
                    Id = 0xABCDEF,
                    InviterIdentity = new Identity { AccountId = Account(3) },
                    InviterName = "Zeratul#77",
                    Program = FourCc.Encode("S2"),
                },
            ],
            SentInvitations = [new BgsSentInvitation { Id = 9, TargetName = "Artanis#5" }],
        };

        var decoded = BgsFriendsSubscribeResponse.Decode(response.Encode());

        Assert.Equal(200u, decoded.MaxFriends);
        Assert.Equal([1ul, 2ul], decoded.Friends.Select(f => f.AccountId.Low));
        Assert.Equal("Kerrigan#1", decoded.Friends[1].BattleTag);
        var invitation = decoded.ReceivedInvitations.Single();
        Assert.Equal(0xABCDEFul, invitation.Id);
        Assert.Equal(3ul, invitation.InviterIdentity!.AccountId!.Low);
        Assert.Equal("Zeratul#77", invitation.InviterName);
        Assert.Equal("S2", FourCc.Decode(invitation.Program!.Value));
        Assert.Equal("Artanis#5", decoded.SentInvitations.Single().TargetName);
    }

    [Theory]
    [InlineData(BgsFriendsNotificationKind.FriendAdded)]
    [InlineData(BgsFriendsNotificationKind.FriendRemoved)]
    [InlineData(BgsFriendsNotificationKind.FriendUpdated)]
    public void FriendNotification_DecodesTargetAndSubscriber(BgsFriendsNotificationKind kind)
    {
        var body = new BgsFriendsNotification
        {
            Kind = kind,
            Friend = new BgsFriendMessage { AccountId = Account(77), BattleTag = "Nova#2" },
            SubscriberAccountId = Account(1),
        }.Encode();

        var decoded = BgsFriendsNotification.Decode((uint)kind, body)!;

        Assert.Equal(kind, decoded.Kind);
        Assert.Equal(77ul, decoded.Friend!.AccountId.Low);
        Assert.Equal("Nova#2", decoded.Friend.BattleTag);
        Assert.Equal(1ul, decoded.SubscriberAccountId!.Low);
    }

    [Fact]
    public void InvitationNotifications_Decode()
    {
        var added = BgsFriendsNotification.Decode(3, new BgsFriendsNotification
        {
            Kind = BgsFriendsNotificationKind.ReceivedInvitationAdded,
            ReceivedInvitation = new BgsReceivedInvitation { Id = 55, InviterName = "Tychus#9" },
            Reason = 0,
        }.Encode())!;
        Assert.Equal(55ul, added.ReceivedInvitation!.Id);
        Assert.Equal(55ul, added.InvitationId);

        var sentRemoved = BgsFriendsNotification.Decode(6, new BgsFriendsNotification
        {
            Kind = BgsFriendsNotificationKind.SentInvitationRemoved,
            SubscriberAccountId = Account(1),
            InvitationId = 66,
            Reason = 2,
        }.Encode())!;
        Assert.Equal(66ul, sentRemoved.InvitationId);
        Assert.Equal(2u, sentRemoved.Reason);

        Assert.Null(BgsFriendsNotification.Decode(99, []));
    }

    [Fact]
    public void PresenceSubscribeRequest_EncodesExpectedBytes()
    {
        var request = new BgsPresenceSubscribeRequest
        {
            EntityId = new EntityId { High = 1, Low = 2 },
            ObjectId = 3,
            Programs = [FourCc.Encode("BN")],
        };

        // entity_id (2) {high fixed64 = 1, low fixed64 = 2}, object_id (3) = 3, program (4, fixed32) = 0x424E.
        var expected = "1212" + "090100000000000000" + "110200000000000000" + "1803" + "254E420000";
        Assert.Equal(expected, Convert.ToHexString(request.Encode()));

        var decoded = BgsPresenceSubscribeRequest.Decode(request.Encode());
        Assert.Equal(2ul, decoded.EntityId.Low);
        Assert.Equal(3ul, decoded.ObjectId);
        Assert.Equal([0x424Eu], decoded.Programs);
    }

    [Fact]
    public void BatchSubscribe_RoundTrips()
    {
        var request = new BgsPresenceBatchSubscribeRequest { EntityIds = [Account(1), Account(2)], Programs = [Bn], ObjectId = 4 };
        var decoded = BgsPresenceBatchSubscribeRequest.Decode(request.Encode());
        Assert.Equal([1ul, 2ul], decoded.EntityIds.Select(e => e.Low));
        Assert.Equal(4ul, decoded.ObjectId);

        var response = BgsPresenceBatchSubscribeResponse.Decode(
            new BgsPresenceBatchSubscribeResponse { Failed = [(Account(2), 12u)] }.Encode());
        Assert.Equal(2ul, response.Failed.Single().EntityId!.Low);
        Assert.Equal(12u, response.Failed.Single().Result);
    }

    [Fact]
    public void FieldOperation_RoundTripsSetAndClear()
    {
        var set = BgsPresenceFieldOperation.Decode(Set(2, 3, new Variant { FourccValue = "S2" }, uniqueId: 7).Encode());
        Assert.Equal(Bn, set.Key.Program);
        Assert.Equal(2u, set.Key.Group);
        Assert.Equal(3u, set.Key.Field);
        Assert.Equal(7ul, set.Key.UniqueId);
        Assert.Equal("S2", set.Value.FourccValue);
        Assert.False(set.Clear);

        Assert.True(BgsPresenceFieldOperation.Decode(Clear(2, 1).Encode()).Clear);
    }

    [Fact]
    public void PresenceListenerNotification_RoundTrips()
    {
        var body = new BgsPresenceListenerNotification
        {
            States = [new BgsPresenceState { EntityId = Account(5), Operations = [Set(1, 4, new Variant { StringValue = "Fenix#3" })] }],
            SubscriberProgram = FourCc.Encode("S2"),
        }.Encode();

        var decoded = BgsPresenceListenerNotification.Decode(body);

        Assert.Equal(5ul, decoded.States.Single().EntityId!.Low);
        Assert.Equal("Fenix#3", decoded.States.Single().Operations.Single().Value.StringValue);
        Assert.Equal(0x5332u, decoded.SubscriberProgram);
    }

    [Fact]
    public void ChannelPresence_ReadsExtension101FromJoinAndUpdate()
    {
        var state = new BgsPresenceState { EntityId = GameAccount(8), Operations = [Set(2, 1, new Variant { BoolValue = true })], Healing = true };

        var join = BgsChannelPresence.Decode(BgsChannelPresence.OnJoinMethod, BgsChannelPresence.Encode(BgsChannelPresence.OnJoinMethod, state))!;
        var update = BgsChannelPresence.Decode(BgsChannelPresence.OnUpdateChannelStateMethod, BgsChannelPresence.Encode(BgsChannelPresence.OnUpdateChannelStateMethod, state))!;

        Assert.Equal(8ul, join.EntityId!.Low);
        Assert.True(join.Healing);
        Assert.True(update.Operations.Single().Value.BoolValue);
        Assert.Null(BgsChannelPresence.Decode(5, BgsChannelPresence.Encode(BgsChannelPresence.OnUpdateChannelStateMethod, state)));
    }

    [Fact]
    public void Tracker_ReadsAccountAndDiscoversGameAccounts()
    {
        var tracker = new BgsPresenceTracker();
        var account = BgsEntityKey.From(Account(10));

        var discovered = tracker.Apply(new BgsPresenceUpdate(account,
        [
            Set(1, BgsPresenceFields.AccountFullName, new Variant { StringValue = "Sarah Kerrigan" }),
            Set(1, BgsPresenceFields.AccountBattleTag, new Variant { StringValue = "Queen#1" }),
            Set(1, BgsPresenceFields.AccountGameAccounts, new Variant { EntityIdValue = GameAccount(100) }, uniqueId: 100),
            Set(1, BgsPresenceFields.AccountGameAccounts, new Variant { EntityIdValue = GameAccount(101) }, uniqueId: 101),
            Set(1, BgsPresenceFields.AccountLastOnline, new Variant { IntValue = 123 }),
        ], IsFullState: true));

        Assert.Equal([100ul, 101ul], discovered.Select(e => e.Low).Order());
        var presence = tracker.GetAccount(account)!;
        Assert.Equal("Sarah Kerrigan", presence.FullName);
        Assert.Equal("Queen#1", presence.BattleTag);
        Assert.Equal(123, presence.LastOnline);
        Assert.Equal(2, presence.GameAccountIds.Count);

        // Seen once, never reported again.
        Assert.Empty(tracker.Apply(new BgsPresenceUpdate(account,
            [Set(1, BgsPresenceFields.AccountGameAccounts, new Variant { EntityIdValue = GameAccount(100) }, uniqueId: 100)], false)));
    }

    [Fact]
    public void Tracker_ReadsGameAccountFields()
    {
        var tracker = new BgsPresenceTracker();
        var game = BgsEntityKey.From(GameAccount(100));
        var rich = new BgsRichPresence(FourCc.Encode("S2"), 0x1234, 17);

        tracker.Apply(new BgsPresenceUpdate(game,
        [
            Set(2, BgsPresenceFields.GameAccountIsOnline, new Variant { BoolValue = true }),
            Set(2, BgsPresenceFields.GameAccountProgram, new Variant { FourccValue = "S2" }),
            Set(2, BgsPresenceFields.GameAccountAwayStatus, new Variant { IntValue = BgsPresenceFields.BusyFlag }),
            Set(2, BgsPresenceFields.GameAccountBattleTag, new Variant { StringValue = "Queen#1" }),
            Set(2, BgsPresenceFields.GameAccountOwner, new Variant { EntityIdValue = Account(10) }),
            Set(2, BgsPresenceFields.GameAccountRichPresence, new Variant { MessageValue = rich.Encode() }),
        ], true));

        var presence = tracker.GetGameAccount(game)!;
        Assert.True(presence.IsOnline);
        Assert.Equal("S2", presence.Program);
        Assert.True(presence.IsBusy);
        Assert.False(presence.IsAway);
        Assert.Equal(10ul, presence.OwnerAccountId!.Value.Low);
        Assert.Equal(rich, presence.RichPresence);
        Assert.Equal("S2", presence.RichPresence!.ProgramName);

        // Owner field links it back without the account's own list.
        Assert.Single(tracker.GetGameAccountsOf(BgsEntityKey.From(Account(10))));

        tracker.Apply(new BgsPresenceUpdate(game, [Clear(2, BgsPresenceFields.GameAccountIsOnline)], false));
        Assert.Null(tracker.GetGameAccount(game)!.IsOnline);
    }

    [Fact]
    public void Tracker_ReadsProgramGivenAsNumber()
    {
        var tracker = new BgsPresenceTracker();
        var game = BgsEntityKey.From(GameAccount(1));
        tracker.Apply(new BgsPresenceUpdate(game, [Set(2, 3, new Variant { UintValue = FourCc.Encode("BSAp") })], true));
        Assert.Equal("BSAp", tracker.GetGameAccount(game)!.Program);
    }

    [Fact]
    public void Directory_CombinesFriendsListWithPresence()
    {
        var directory = new BgsFriendsDirectory();
        directory.Apply(new BgsFriendsSubscribeResponse
        {
            Friends = [new BgsFriendMessage { AccountId = Account(10) }, new BgsFriendMessage { AccountId = Account(20), BattleTag = "Swann#4" }],
        });

        directory.Apply(new BgsPresenceUpdate(BgsEntityKey.From(Account(10)),
        [
            Set(1, BgsPresenceFields.AccountBattleTag, new Variant { StringValue = "Queen#1" }),
            Set(1, BgsPresenceFields.AccountFullName, new Variant { StringValue = "Sarah Kerrigan" }),
            Set(1, BgsPresenceFields.AccountGameAccounts, new Variant { EntityIdValue = GameAccount(100) }, 100),
            Set(1, BgsPresenceFields.AccountGameAccounts, new Variant { EntityIdValue = GameAccount(200) }, 200),
        ], true));
        directory.Apply(new BgsPresenceUpdate(BgsEntityKey.From(GameAccount(100)),
            [Set(2, 1, new Variant { BoolValue = true }), Set(2, 3, new Variant { FourccValue = "BSAp" })], true));
        directory.Apply(new BgsPresenceUpdate(BgsEntityKey.From(GameAccount(200)),
            [Set(2, 1, new Variant { BoolValue = true }), Set(2, 3, new Variant { FourccValue = "S1" }), Set(2, 10, new Variant { BoolValue = true })], true));

        var friends = directory.GetFriends();

        var kerrigan = friends[0];
        Assert.Equal("Queen#1", kerrigan.BattleTag);
        Assert.Equal("Sarah Kerrigan", kerrigan.FullName);
        Assert.True(kerrigan.IsOnline);
        Assert.Equal("S1", kerrigan.Program); // a game wins over the Battle.net app
        Assert.True(kerrigan.IsAway);

        var swann = friends[1];
        Assert.Equal("Swann#4", swann.BattleTag);
        Assert.False(swann.IsOnline);
        Assert.Null(swann.Program);

        // Signing out of the game leaves only the app.
        directory.Apply(new BgsPresenceUpdate(BgsEntityKey.From(GameAccount(200)), [Set(2, 1, new Variant { BoolValue = false })], false));
        Assert.Equal("BSAp", directory.GetFriend(BgsEntityKey.From(Account(10)))!.Program);

        directory.Apply(new BgsFriendsNotification
        {
            Kind = BgsFriendsNotificationKind.FriendRemoved,
            Friend = new BgsFriendMessage { AccountId = Account(20) },
        });
        Assert.Single(directory.GetFriends());
    }

    [Fact]
    public void FrontClient_Dispatch_AppliesListenerCallsToSocial()
    {
        var client = new FrontClient();
        var notified = 0;
        var changed = 0;
        client.FriendsNotified += _ => notified++;
        client.SocialChanged += () => changed++;

        client.Dispatch(
            new Header { ServiceId = 0, Token = 1, MethodId = 1, ServiceHash = ServiceHash.Compute(FrontServices.FriendsListener) },
            new BgsFriendsNotification { Kind = BgsFriendsNotificationKind.FriendAdded, Friend = new BgsFriendMessage { AccountId = Account(10) } }.Encode());

        client.Dispatch(
            new Header { ServiceId = 0, Token = 2, MethodId = 2, ServiceHash = ServiceHash.Compute(FrontServices.PresenceListener) },
            new BgsPresenceListenerNotification
            {
                States =
                [
                    new BgsPresenceState
                    {
                        EntityId = Account(10),
                        Operations = [Set(1, 4, new Variant { StringValue = "Queen#1" }), Set(1, 3, new Variant { EntityIdValue = GameAccount(100) }, 100)],
                    },
                ],
            }.Encode());

        // Older path: ChannelListener.OnUpdateChannelState, with the client/server flag bit set on the method id.
        client.Dispatch(
            new Header { ServiceId = 0, Token = 3, MethodId = 6 | 0x40000000, ServiceHash = ServiceHash.Compute(FrontServices.ChannelListener) },
            BgsChannelPresence.Encode(6, new BgsPresenceState
            {
                EntityId = GameAccount(100),
                Operations = [Set(2, 1, new Variant { BoolValue = true }), Set(2, 3, new Variant { FourccValue = "S2" })],
            }));

        var friend = client.Social.GetFriends().Single();
        Assert.Equal("Queen#1", friend.BattleTag);
        Assert.True(friend.IsOnline);
        Assert.Equal("S2", friend.Program);
        Assert.Equal(1, notified);
        Assert.Equal(3, changed);
    }

    [Fact]
    public void FrontClient_Dispatch_ReportsUnknownAndUnreadableCallsWithoutThrowing()
    {
        var client = new FrontClient();
        Header? unhandled = null;
        Exception? error = null;
        client.UnhandledCall += (h, _) => unhandled = h;
        client.SocialError += ex => error = ex;

        client.Dispatch(new Header { ServiceId = 0, Token = 1, MethodId = 9, ServiceHash = 0x12345678 }, []);
        Assert.Equal(0x12345678u, unhandled!.ServiceHash);

        // Truncated: claims a 16-byte field but carries one byte.
        client.Dispatch(new Header { ServiceId = 0, Token = 2, MethodId = 1, ServiceHash = ServiceHash.Compute(FrontServices.FriendsListener) }, [0x0A, 0x10, 0x00]);
        Assert.NotNull(error);
    }
}
