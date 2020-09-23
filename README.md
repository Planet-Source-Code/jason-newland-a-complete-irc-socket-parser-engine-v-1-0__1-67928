<div align="center">

## A Complete IRC Socket Parser Engine v\.1\.0


</div>

### Description

This is a complete IRC socket and parser engine suitable for simply dropping in your project. It contains the socket API, socket control commands (similar to normal winsock properteries .connect, etc.. see Read Me for more details). It will also parse out the raw socket data, split it into its relevant parts, raise an event that the data has parsed, which can then be sent on to the second parser. The second parser splits the data into events, such as TEXTCHAN, ACTIONCHAN, JOIN, PART, etc and raises its associated event. All you need to simply do is supply the code to tell your client what to do after that event has been raised. Makes making an IRC client a lot simpler in the long run, as most of the hard work (the parser) has been done for you.

Please enjoy this, a don't forget to vote and comment on my code :)
 
### More Info
 
See code

Requires an understanding of declaring and using class modules as an object. IE: Public WithEvents sParser As cIRCParser, Set sParser = New cIRCParser, etc

See Read ME


<span>             |<span>
---                |---
**Submitted On**   |2006-11-22 10:53:32
**By**             |[Jason Newland](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jason-newland.md)
**Level**          |Advanced
**User Rating**    |5.0 (35 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[A\_Complete2049172212007\.zip](https://github.com/Planet-Source-Code/jason-newland-a-complete-irc-socket-parser-engine-v-1-0__1-67928/archive/master.zip)








