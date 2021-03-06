# shboom
Tom Torok's original ASP app for searching database tables

"I've never found anything easier than Tom Torok's stuff -- all you have to do is point it to a SQL Server table, specify the columns you want to search, any column headers you want to see in a different form than the field names, and voila."

-- Tim Henderson, August 23, 2005, on the NICAR-L listserv.

### Introduction

At one of the NICAR conferences I went to in the early 2000s, Tom Torok of The New York Times (and previously The Philadelphia Inquirer) showed off a tool he had written that would make your SQL Server database searchable. Any database, no matter the type of information contained in it. He called it "Shboom."

Sure, there was a bit of configuration to make the database, as my former colleague Aron Pilhofer put it, "shboomable," and it was written in ASP, but hey, this was pre-Django and pre-Rails. For most of us, it was pre-PHP. Shboom was, like most everything Tom did, almost magical and slightly mad. I figured it needed to be remembered, because it was a remarkable piece of software for those of us who had no idea what software really was. It is one of the first examples of web development in the newsroom that spread beyond it.

So here, then, are files that, with Tom's permission, I pulled off a New York Times server in 2014 for the purposes of open-sourcing Shboom. The design of Shboom's pages is by Matt Ericson, also of The New York Times.

### FAQ

##### Does it still work?

Yes! From Tom: "I've recently written a number of Shbooms for The Philadelphia Inquirer. It's web based and I used the scripting engine for whatever is associated with Windows Explorer in Windows 7. The scripting engine of some versions of Firefox will take you to the end and produce nothing. That applies to writing pages only. The pages it writes work in all browsers for searches and more (more on that later). The latest version of MS SQL Server I tried it on was 2012 and it worked."

##### What do I need to find out?

MS SQL Server, probably something like IIS and ASP. Which you probably have CDs for somewhere.

##### Why are you doing this?

Because Shboom deserves to be remembered for what it was and the role it played in getting people involved in building software for newsrooms. Because so much of the early newsroom web development work has been lost or forgotten. Because as the first rough draft of history, journalism has a duty to preserve important things.

Also, as a way of saying thank you to a friend and colleague who was generous with his time and expertise.
