<%

' Following are all the configuration items with default values set.
' Override them if you want in a separate file, see e.g. /web1/ow.asp.

' "The Truth about MS Access" : http://www.15seconds.com/Issue/010514.htm
' OPENWIKI_DB = "Driver={Microsoft Access Driver (*.mdb)};DBQ=e:\data\openwiki\OpenWiki.mdb"
' OPENWIKI_DB = "Driver={SQL Server};server=mymachine;uid=openwiki;pwd=openwiki;database=OpenWiki"
' OPENWIKI_DB = "Driver={Microsoft ODBC for Oracle};Server=OW;Uid=lawrence;Pwd=aphex2twin;"
' OPENWIKI_DB = "MySystemDSName"
' OPENWIKI_DB = "MySQLOpenWiki"
' OPENWIKI_DB = "PostgreSQLOpenWiki"
OPENWIKI_DB = "OpenWikiDist"

'OPENWIKI_DB_SYNTAX = DB_ACCESS               ' see owpreamble.asp for possible values

OPENWIKI_IMAGEPATH       = "ow/images"        ' path to images directory
OPENWIKI_ICONPATH        = "ow/images/icons"  ' path to icons directory
OPENWIKI_ENCODING        = "ISO-8859-1"       ' character encoding to use
OPENWIKI_TITLE           = "O p e n W i k i"  ' title of your wiki
OPENWIKI_FRONTPAGE       = "FrontPage"        ' name of your front page.
OPENWIKI_SCRIPTNAME      = "ow.asp"           ' "mydir/ow.asp" : in case the auto-detected scriptname isn't correct
OPENWIKI_STYLESHEETS     = "ow/xsl/"          ' the subdirectory where the stylesheet files (*.xsl) are located
OPENWIKI_MAXTEXT         = 204800             ' Maximum 200K texts
OPENWIKI_MAXINCLUDELEVEL = 5                  ' Maximum depth of Include's
OPENWIKI_RCNAME          = "RecentChanges"    ' Name of recent changes page (change space to _)
OPENWIKI_RCDAYS          = 30                 ' Default number of RecentChanges days
OPENWIKI_MAXTRAIL        = 5                  ' Maximum number of links in the trail
OPENWIKI_STOPWORDS       = "StopWords"        ' Name of page containing stop words (change space to _). Stop words are words that won't be hyperlinked. Use empty string "" if you do not want to support stop words.
OPENWIKI_TEMPLATES       = "Template$"        ' Pattern for templates usable when creating a new page
OPENWIKI_TIMEZONE        = "+01:00"           ' Timezone of the server running this wiki, valid values are e.g. "+04:00", "-09:00", etc.
OPENWIKI_MAXNROFAGGR     = 150                ' Maximum number of rows to show in an aggregated feed
OPENWIKI_MAXWEBGETS      = 3                  ' Maximum number of RSS feeds that may be refreshed from a remote server for one user request.
OPENWIKI_SCRIPTTIMEOUT   = 120                ' Maximum amount of seconds to wait for RSS feeds to be syndicated, if set to 0 the default timeout value of ASP is used.
OPENWIKI_DAYSTOKEEP      = 30                 ' Number of days to keep old revisions
OPENWIKI_UPLOADDIR       = "attachments/"     ' The virtual directory where uploads are stored
OPENWIKI_MAXUPLOADSIZE   = 8388608            ' Use to limit the size of uploads, in bytes (default = 8,388,608)
OPENWIKI_UPLOADTIMEOUT   = 300                ' Timeout in seconds (upload must succeed within this time limit)


MSXML_VERSION = 3   ' use 4 if you've installed MSXML v4.0

gReadPassword = ""    ' use empty string "" if anyone may read
gEditPassword = ""    ' use empty string "" if anyone may edit
' In case you want more sophisticated security, then you should
' rely on the Integrated Windows authentication feature of IIS.

gDefaultBookmarks = "FrontPage RecentChanges TitleIndex UserPreferences RandomPage Help"

' Major system options
cReadOnly             = 0        ' 1 = readonly wiki         0 = editable wiki
cNakedView            = 0        ' 1 = run in naked mode     0 = show headers/footers
cUseSubpage           = 1        ' 1 = use /subpages         0 = do not use /subpages
cFreeLinks            = 1        ' 1 = use [[word]] links    0 = LinkPattern only
cWikiLinks            = 1        ' 1 = use LinkPattern       0 = possibly allow [[word]] only
cAcronymLinks         = 0        ' 1 = link acronyms         0 = do not link 3 or more capitalized characters
cTemplateLinking      = 1        ' 1 = allow TemplateName->WikiLink   0 = don't do template linking
cRawHtml              = 0        ' 1 = allow <html> tag      0 = no raw HTML in pages
cMathML               = 0        ' 1 = allow <math> tag      0 = no raw math in pages
cHtmlTags             = 0        ' 1 = "unsafe" HTML tags    0 = only minimal tags
cCacheXSL             = 0        ' 1 = cache stylesheet      0 = don't cache stylesheet
cCacheXML             = 0        ' 1 = cache partial results 0 = do not cache partial results
cAllowRSSExport       = 1        ' 1 = allow RSS feed        0 = do not export your pages to RSS
cAllowNewSyndications = 1        ' 1 = allow new URLs to be syndicated    0 = only allow syndication of the URLs in the database table openwiki_rss
cAllowAggregations    = 1        ' 1 = allow aggregation of syndications (note: you MUST use MSXML v3 sp2 for this to work)   0 = do not allow aggregrations
cEmbeddedMode         = 0        ' 1 = embed the wiki into another app    0 = process browser request
cAllowAttachments     = 0        ' 1 = allow attachments     0 = do not allow attachments (WARNING: Allowing attachments poses a security risk!! See file owattach.asp)

' Minor system options
cSimpleLinks          = 1        ' 1 = only letters,         0 = allow _ and numbers
cNonEnglish           = 1        ' 1 = extra link chars,     0 = only A-Za-z chars
cNetworkFile          = 0        ' 1 = allow remote file:    0 = no file:// links
cBracketText          = 1        ' 1 = allow [URL text]      0 = no link descriptions
cBracketIndex         = 1        ' 1 = [URL] -> [<index>]    0 = [URL] -> [URL]
cHtmlLinks            = 1        ' 1 = allow A HREF links    0 = no raw HTML links
cBracketWiki          = 1        ' 1 = [WikiLnk txt] link    0 = no local descriptions
cShowBrackets         = 0        ' 1 = keep brackets         0 = remove brackets when it's an external link
cFreeUpper            = 1        ' 1 = force upper case      0 = do not force case for free links
cLinkImages           = 1        ' 1 = display image         0 = display link to image
cUseHeadings          = 1        ' 1 = allow = h1 text =     0 = no header formatting
cUseLookup            = 1        ' 1 = lookup host names     0 = skip lookup (IP only)
cStripNTDomain        = 1        ' 1 = strip NT domainname   0 = keep NT domainname in remote username
cMaskIPAddress        = 1        ' 1 = mask last part of IP  0 = show full IP address in RecentChanges list, etc.
cOldSkool             = 1        ' 1 = use '' and '''        0 = don't use '' and ''' for italic and bold, and use Wiki''''''Link to escape WikiLink
cNewSkool             = 1        ' 1 = use //, **, -- and __ 0 = don't use //, **, -- and __ for italic, bold, strikethrough and underline and use ~WikiLink to escape WikiLink
cNumTOC               = 0        ' 1 = TOC numbered          0 = TOC just indented text
cNTAuthentication     = 1        ' 1 = Use NT username       0 = blank username in preferences
cDirectEdit           = 1        ' 1 = go direct to edit     0 = go to blank page first
cAllowCharRefs        = 1        ' 1 = allow char refs       0 = no character references allowed (like &copy; or &#151;)
cWikifyHeaders        = 0        ' 1 = wikify headers        0 = do not apply wiki formatting within headers

' User options
cEmoticons            = 1        ' 1 = use emoticons         0 = don't show feelings
cUseLinkIcons         = 0        ' 1 = icons for ext links   0 = no icon images for external links
cPrettyLinks          = 1        ' 1 = display Words Smashed Together     0 = display WordsSmashedTogether
cExternalOut          = 1        ' 1 = external links open in new window, 0 = open in same window

%>