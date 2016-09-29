# Custom Lync Client

The application initializes the platform and a user endpoint and subscribes to the sample user's presence (self-presence). The sample then publishes the user state, machine state, note, and contact card. The note is published using raw XML, whereas all other categories are published using strongly-typed objects.
The client can delete the presence categories that it published, terminates the endpoint, and shuts down the platform, exiting normally.

You may log in to the same user as the sample using a client (such as Microsoft Lync), to see the categories being published as well.

# Features
- Presence Publication using the grammar and strongly-typed categories
- Capabilities publication
- Tray icon support
- Status change message notifications
- Code console output redirection

# Prerequisites
- 64 bit system
- Unified Communications Managed API (UCMA) [Runtime](https://www.microsoft.com/en-us/download/details.aspx?id=47344) or [SDK](https://www.microsoft.com/en-us/download/details.aspx?id=47345)
- Microsoft Lync Server
- One user, enabled to use the Microsoft Lync Server
- The credentials for that user (or kerberos ticket :) 
- Visual Studio

# Running Custom Lync Client
1. You may want to supply the configuration settings to be used by the Custom Lync Client in the accompanying App.config file.
2. Open the project in Visual Studio, and hit F5.
3. Change the presence of the user logged in on Microsoft Lync, and see the presence change respectively.

# Information
- [Lync endpoint types](https://msdn.microsoft.com/en-us/library/office/dn345988.aspx)

# License
Feel free to improve the Custom Lync Client software with any functionality by introducing changes.

Commercial use prohibited.
