
R E A D M E
===========


License
-------
See file LICENSE.txt.


FAQ
---
Q1. Can I use the source code in any way I like?
A1. That's right.

Q2. Can I hold you responsible if it doesn't do what I expect it to do?
A2. Nope.

Q3. Fair enough, thanks.
A3. No worries.


Installation
------------
Before OpenWiki can run you must have the following components installed:
    
    o  MS ADO.DB v2.5 or higher
    o  VBScript v5.5 or higher
    o  MSXML v3 SP2 or higher
    o  A Web Server (i.e. PWS or IIS)
    o  A database like MS Access, SQL-Server, Oracle or MySQL

Please read the file INSTALL.txt for further instructions on how to finish
the installation.


Attachments
-----------
Please read the comments in file ow/owattach.asp for more information on how
to support attachments cq file uploads in your wiki.

This software should be released with the variable cAllowAttachments set to 
value 0 in the file ow/owconfig_default.asp. If it's not 0, and you have 
not read the comments in file ow/owattach.asp set it to 0 immediately ! 
Be very paranoid. ;-)


Integration into existing applications
--------------------------------------
Even when you do not plan to run a wiki site, you might want to integrate
this software into one or more of your existing applications. Run and see 
the code of control.asp, and imagine what you can do.

Where I work we've integrated OpenWiki into all our internal applications.
The applications are using very structured data, but where one can input
text one can use all the flexibility and ease of the OpenWiki formatting 
rules to make the text look "good".


Aggregating RSS feeds
---------------------
You MUST use the MSXML v3 sp2 component from Microsoft if you want to use 
the <Aggregate> macro. It does not work with MSXML v4.


Deleting pages and attachments
------------------------------
Users of an OpenWiki site can not delete pages and attachments. At best 
they can mark pages and attachments as "deprecated". See also the help 
page at http://openwiki.com/?HelpOnProcessingInstructions#deprecated.

Once a page or attachment is marked as "deprecated" it is not deleted 
until the OpenWiki administrator (that is you) runs the script 
deprecate.asp. 

You should place this script at a location within your website that only 
you have access to (see IIS help to find out how to do that). Or simply 
rename to script to something arcane that only you can remember. 

Open the script in your fav texteditor. You will see lines:

    Response.Write("Look in the script! <p />")
    
    ' RUN AT YOUR OWN RISK !!!
    ' ALSO DELETES DEPRECATED ATTACHMENTS
    ' COMMENT THE NEXT LINE AND REFRESH THE PAGE
    Response.End

    ' If a page is marked as deprecated, but was last modified less than
    ' <gDaysToKeep> days, then keep the page. Otherwise delete it.
    gDaysToKeep = 30

Comment the lines that start with "Response.". In essence you'd get:

    ' Response.Write("Look in the script! <p />")
    ' Response.End
    gDaysToKeep = 30

Then modify the variable gDaysToKeep to your liking. Lastly, run the script
to destroy deprecated pages and attachments.


Note to developers
------------------
When you develop your own stuff, it helps to set the system options cCacheXSL 
and cCacheXML to value 0. ;-)

Also, at various places you might want to uncomment the line:

     On Error Resume Next
     
most notably in method owmacros.ExecMacro and sometimes in 
owtransformer.LoadXSL when you're changing the xsl stylesheets.

Those who used OpenWiki prior to v0.78 will notice a big difference in
the structure of the program: instead of 1 big asp file it's now split
into several asp files. The idea is that it should become easier for
developers to add their stuff and yet be very easy to upgrade to future 
versions of OpenWiki. Unfortunately, this release hasn't been finished 
yet in that regard. Hopefully v0.79 will have this feature.

Output should also always be XHTML v1.0 compliant. Which means the output 
is always a well-formed and valid XML document. It also means that it is 
impossible to enter invalid wiki formats. To view the XML output before it
passes through the XSLT engine add &xml=1 to the end of the URL. 
Example: http://openwiki.com/?FrontPage&xml=1

From an architectural point of view: the output by the wiki engine is in 
XML format. An XSL stylesheet is used to transform the XML output into 
XHTML output. It would be quite easy to transform the XML output to other 
formats such as WML, limited HTML, etc.

This software is also usable in non-wiki environments. It can be used to 
create a somewhat more sophicated textarea control that allows the ease of 
formatting like wiki's do. See file control.asp for an example.

There's an abstraction layer for all data manupulation and retrieval 
functions. Currently an implementation exists for databases, supported 
are MS Access, MS SQL-Server, Oracle and mySQL. I haven't tested on
PostgreSQL yet, but if someone wants support for that, no problem, I'll
do it.

I don't plan to write an implementation that uses the filesystem; why would
you, when you can rely on databases to do that work for you? On Windows
machines there's at least always (the drivers of) MS Access available.

All character sets should work. But this is not tested extensively as 
of yet. Still, I've gotten no complaints about it either. ;)

User authorization support
- via basic authentication handled by webserver
- via form based authentication


Bugs, suggestions, comments
---------------------------
Please report installation issues at http://openwiki.com/?/Installation
Please report bugs at http://openwiki.com/?/Bugs
Please enter suggestions at http://openwiki.com/?/Suggestions

Or alternatively send them and any other comments you may have via email
to lawrence.pit@openwiki.com.


Thank you
---------
Thanks for the inspiration to:
Clifford Adams of UseMod (http://www.usemod.com),
Jürgen Hermann of Moin Moin (http://moin.sourceforge.net/cgi-bin/moin/moin/)
Ward Cunningham of the original Wiki (http://www.c2.com).



Lawrence Pit
(c) 2000 - 2002 OpenWiki
http://openwiki.com

