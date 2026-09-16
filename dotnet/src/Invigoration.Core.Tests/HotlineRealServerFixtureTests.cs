using System.Buffers.Binary;
using System.Reflection;
using Invigoration.Core.Hotline;

namespace Invigoration.Core.Tests;

/// <summary>
/// Bytes captured from a real server (MacDomain, as a guest, 2026-09-16) rather than bytes this
/// code made up. Two bugs got through a full set of self-built fixtures because the fixtures
/// encoded the same wrong layout the parser decoded — these are the antidote.
/// </summary>
public class HotlineRealServerFixtureTests
{
    /// <summary>A real FileNameWithInfo row: a text file whose name begins with spaces, which is what exposed the offset bug.</summary>
    private const string RealTextFileEntry =
        "54455854747478740000000f000000000000002c2020202020736565204e45575320666f7220496e666f2061626f75742046696c65732020202020202e747874";

    /// <summary>A real folder row.</summary>
    private const string RealFolderEntry =
        "666c647200000000000000020000000000000015202020666f7220796f757220646f776e6c6f616473";

    /// <summary>MacDomain's actual access bitmap for a guest account.</summary>
    private const string RealGuestAccessBitmap = "60700c2003800000";

    [Fact]
    public void AFileRowFromARealServer_ParsesWithItsNameIntact()
    {
        var entry = HotlineFileEntry.TryParse(Convert.FromHexString(RealTextFileEntry));

        Assert.NotNull(entry);
        Assert.Equal("TEXT", entry.TypeCode);
        Assert.False(entry.IsFolder);
        Assert.Equal(15u, entry.Size);

        // Leading spaces are the server's own formatting — kept, not trimmed, and crucially not
        // mistaken for the name's length.
        Assert.Equal("     see NEWS for Info about Files      .txt", entry.Name);
    }

    [Fact]
    public void AFolderRowFromARealServer_Parses()
    {
        var entry = HotlineFileEntry.TryParse(Convert.FromHexString(RealFolderEntry));

        Assert.NotNull(entry);
        Assert.True(entry.IsFolder);
        Assert.Equal(2u, entry.Size); // items inside, not bytes
        Assert.Equal("   for your downloads", entry.Name);
    }

    /// <summary>
    /// The exact shape of the bug: a name starting with two spaces reads as a length of 0x2020
    /// when taken from the wrong offset, which fails the bounds check and drops the row — so an
    /// entire listing silently became "no files".
    /// </summary>
    [Fact]
    public void EveryRowOfARealListing_Survives()
    {
        foreach (var hex in new[] { RealTextFileEntry, RealFolderEntry })
        {
            Assert.NotNull(HotlineFileEntry.TryParse(Convert.FromHexString(hex)));
        }
    }

    private static HotlineTransactionClient ClientWithAccessBits(string hex)
    {
        var client = new HotlineTransactionClient();
        typeof(HotlineTransactionClient)
            .GetProperty(nameof(HotlineTransactionClient.OwnAccessBits))!
            .SetValue(client, BinaryPrimitives.ReadUInt64BigEndian(Convert.FromHexString(hex)));
        return client;
    }

    /// <summary>
    /// A guest on MacDomain can download, upload, read and post news. Reading the bitmap the other
    /// way round denied all four — which is what produced "this account isn't allowed to browse
    /// files" for an account that plainly was.
    /// </summary>
    [Fact]
    public async Task ARealGuestAccount_HasThePrivilegesTheServerActuallyGrantedIt()
    {
        await using var client = ClientWithAccessBits(RealGuestAccessBitmap);

        Assert.True(client.CanDownloadFiles, "a guest on MacDomain can browse and download");
        Assert.True(client.CanUploadFiles);
        Assert.True(client.CanReadNews);
        Assert.True(client.CanPostNews);
    }

    /// <summary>The same bitmap must NOT grant what it doesn't — a check that the bit order isn't just flipped into being permissive.</summary>
    [Fact]
    public async Task ARealGuestAccount_DoesNotHaveAdministrativePrivileges()
    {
        await using var client = ClientWithAccessBits(RealGuestAccessBitmap);

        Assert.False(client.HasOwnAccess(HotlineAccessBits.CreateUser));
        Assert.False(client.HasOwnAccess(HotlineAccessBits.DeleteUser));
        Assert.False(client.HasOwnAccess(HotlineAccessBits.DisconnectUser));
        Assert.False(client.HasOwnAccess(HotlineAccessBits.CannotBeDisconnected));
        Assert.False(client.HasOwnAccess(HotlineAccessBits.DeleteFile));
    }

    /// <summary>Bit 0 is the top bit of the first byte — the single fact the protocol docs never state.</summary>
    [Theory]
    [InlineData("8000000000000000", 0)]
    [InlineData("4000000000000000", 1)]
    [InlineData("2000000000000000", 2)]
    [InlineData("0000000000000001", 63)]
    public async Task BitNumbering_RunsFromTheTopBitOfTheFirstByte(string hex, int expectedBit)
    {
        await using var client = ClientWithAccessBits(hex);

        for (var bit = 0; bit < 64; bit++)
        {
            Assert.Equal(bit == expectedBit, client.HasOwnAccess(bit));
        }
    }
}
