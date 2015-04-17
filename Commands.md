# Command Structure #

Because Invigoration is JUST a chat bot, these commands are either internal ( using /) or external using the trigger set and only used by the 'Master'.


# Command List #

idle : Example: /idle 56 uptime (Bot will spit out uptime every 56 seconds)
  * off -- Turns idle off
  * uptime -- Displays uptime of bot in months/weeks/days/hours/minutes/seconds **CURRENTLY DOESN'T WORK!**
  * ver -- Displays current version of the bot.
  * Else... Whatever else you put here, is the idle msg.


disconnect : Example: /disconnect (Alt: disc)
  * bot closes wsbnet and wsbnls connections.

reconnect : Example: /reconnect
  * bot closes wsbnet and wsbnls connections and reconnects to battle.net.

colors : Example: /colors (Alt: color)
  * bot displays information for using color codes with invig. These are INVIG Only!!

hex : Example: /hex This is hex chat! (Alt: h)
  * bot encrypts text as HEX

invigencrypt : Example: /invigencrypt This is invig chat! (alt: i, ie, encrypt)
  * Bot sends text as encrypted in invig encrypt.

sysinfo : Example: /sysinfo
  * Bot displays system info on system its running on.

ver : Example: /ver
  * Bot displays version of itself.

uptime : Example: /uptime
  * Bot displays current uptime (Connected time!)

about : Example: /about
  * Bot displays information about itself.

say :  Example: !say Cheese!
  * Bot will send whatever text you tell it to.

bancount : Example: /bancount
  * Bot displays count of bans since it entered the channel (counts double bans each!)

kickcount : Example: /kickcount
  * Bot displays count of KICKS since it entered the channel (Counts double kicks!)

joincoint : Example: /joincount
  * Bot displays number of joins its seen since it entered the channel. (resets upon leaving!)

user : Example: /user Tagban
  * prepends the words: "Tagban:" to the beginning of EVERYTHING. This is called user focusing. This can also be used via right clicking the username.

useroff : Example: /useroff
  * Ends user focus.

prepend : Example: "/prepend /me is a vagina that " then type: 'smells weird?'
  * will output: Tagban is a vagina that smells weird.