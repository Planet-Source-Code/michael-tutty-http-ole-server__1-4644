<div align="center">

## HTTP OLE Server


</div>

### Description

This code provides the basic framework for HTTP services. There is NO functionality built in to transfer files - that's really not the point of this DLL. If you want to create an HTTP service (for remote function calls, database access, customized web server) you can use this DLL to do it.

What's the differenece between this and ASP? Two main things:

1. You can program this server to run on any port you'd like, without disturbing "normal" web services (PWS, IIS, etc.)

2. (Win95/98 only) You can shut this down when you want to (can't forcibly unload ASP's under 95/98, so debugging DLL's becomes quite a pain).
 
### More Info
 
All you need to know is:

1. Reference the DLL (or pull the source code into your project),

2. Create a class that has a WithEvents variable of type HTTPServer,

3. In the OnRequest event, write the code to respond to the requestor (via the HTTPRequest object)


<span>             |<span>
---                |---
**Submitted On**   |1999-11-07 17:05:22
**By**             |[Michael Tutty](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-tutty.md)
**Level**          |Unknown
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD2001\.zip](https://github.com/Planet-Source-Code/michael-tutty-http-ole-server__1-4644/archive/master.zip)








