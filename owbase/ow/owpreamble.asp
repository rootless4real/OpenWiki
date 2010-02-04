<%

Const OPENWIKI_VERSION     = "0.78"
Const OPENWIKI_REVISION    = "$Revision: 1.2 $"
Const OPENWIKI_XMLVERSION  = "0.91"
Const OPENWIKI_NAMESPACE   = "http://openwiki.com/2001/OW/Wiki"

' possible values for OPENWIKI_DB_SYNTAX
Const DB_ACCESS      = 0
Const DB_SQLSERVER   = 0
Const DB_ORACLE      = 1
Const DB_MYSQL       = 2
Const DB_POSTGRESQL  = 3

' declare 'constants'
Dim OPENWIKI_DB, OPENWIKI_DB_SYNTAX
Dim OPENWIKI_ICONPATH, OPENWIKI_IMAGEPATH, OPENWIKI_ENCODING, OPENWIKI_TITLE, OPENWIKI_FRONTPAGE
Dim OPENWIKI_SCRIPTNAME, OPENWIKI_STYLESHEETS, OPENWIKI_MAXTEXT, OPENWIKI_MAXINCLUDELEVEL
Dim OPENWIKI_RCNAME, OPENWIKI_RCDAYS, OPENWIKI_MAXTRAIL, OPENWIKI_TEMPLATES
Dim OPENWIKI_TIMEZONE, OPENWIKI_MAXNROFAGGR, OPENWIKI_MAXWEBGETS, OPENWIKI_SCRIPTTIMEOUT
Dim OPENWIKI_DAYSTOKEEP
Dim OPENWIKI_STOPWORDS
Dim OPENWIKI_UPLOADDIR, OPENWIKI_MAXUPLOADSIZE, OPENWIKI_UPLOADTIMEOUT

Dim MSXML_VERSION

' declare options
Dim gReadPassword, gEditPassword, gDefaultBookmarks
Dim cReadOnly, cNakedView, cUseSubpage, cFreeLinks, cWikiLinks, cAcronymLinks, cTemplateLinking, cRawHtml, cMathML, cHtmlTags, cCacheXSL, cCacheXML, cDirectEdit, cEmbeddedMode
Dim cSimpleLinks, cNonEnglish, cNetworkFile, cBracketText, cBracketIndex, cHtmlLinks, cBracketWiki, cFreeUpper, cLinkImages, cUseHeadings, cUseLookup, cStripNTDomain, cMaskIPAddress, cOldSkool, cNewSkool, cNumTOC, cAllowCharRefs, cWikifyHeaders
Dim cEmoticons, cUseLinkIcons, cPrettyLinks, cExternalOut
Dim cAllowRSSExport, cAllowNewSyndications, cAllowAggregations, cNTAuthentication, cShowBrackets
Dim cAllowAttachments

' global variables
Dim gLinkPattern, gSubpagePattern, gStopWords, gTimestampPattern, gUrlProtocols, gUrlPattern, gMailPattern, gInterSitePattern, gInterLinkPattern, gFreeLinkPattern, gImageExtensions, gImagePattern, gDocExtensions, gNotAcceptedExtensions, gISBNPattern, gHeaderPattern, gMacros
Dim gFS, gIndentLimit
gFS = Chr(179)           ' The FS character is a superscript "3"
gIndentLimit = 20        ' maximum indent level for bulleted/numbered items

' incoming parameters
Dim gPage                ' page to be worked on
Dim gRevision            ' revision of page to be worked on
Dim gAction              ' action
Dim gTxt                 ' text value passed to input boxes

Dim gServerRoot          ' URL path to script
Dim gScriptName          ' Name of this script
Dim gTransformer         ' transformer of XML data
Dim gNamespace           ' namespace data
Dim gRaw                 ' vector or raw data used by Wikify function
Dim gBracketIndices      ' keep track of the bracketed indices
Dim gTOC                 ' table of contents
Dim gIncludeLevel        ' recursive level of included pages
Dim gCurrentWorkingPages ' stack of pages currently working on when including pages
Dim gIncludingAsTemplate ' including subpages as template
Dim gNrOfRSSRetrievals   ' nr of remote calls performed to retrieve an RSS feed
Dim gAggregateURLs       ' URL's to RSS feeds that need to be aggregated for this page
Dim gCookieHash          ' Hash value to use in cookie names
Dim gTemp                ' temporary value that may be used at all times
Dim gActionReturn        ' return value used by actions
Dim gMacroReturn         ' return value used by macros

If (ScriptEngineMajorVersion < 5) Or (ScriptEngineMajorVersion = 5 And ScriptEngineMinorVersion < 5) Then
    Response.Write("<h2>Error: Missing VBScript v5.5</h2>")
    Response.Write("In order for this script to work correctly the component " _
                 & "VBScript v5.5 " _
                 & "or a higher version needs to be installed on the server. You can download this component from " _
                 & "<a href=""http://msdn.microsoft.com/scripting/"">http://msdn.microsoft.com/scripting/</a>.")
    Response.End
End If

'Dim c, i
'c = Request.ServerVariables.Count
'For i = 1 To c
'    Response.Write(Request.ServerVariables.Key(i) & " ==> " & Request.ServerVariables.Item(i) & "<br>")
'Next
'Response.End

%>