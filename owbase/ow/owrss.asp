<%
'
' ---------------------------------------------------------------------------
' Copyright(c) 2000-2002, Lawrence Pit
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions
' are met:
'
'   * Redistributions of source code must retain the above copyright
'     notice, this list of conditions and the following disclaimer.
'   * Redistributions in binary form must reproduce the above
'     copyright notice, this list of conditions and the following
'     disclaimer in the documentation and/or other materials provided
'     with the distribution.
'   * Neither the name of OpenWiki nor the names of its contributors
'     may be used to endorse or promote products derived from this
'     software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
' "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
' LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
' FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE
' REGENTS OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
' INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
' BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
' CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT
' LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN
' ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
' POSSIBILITY OF SUCH DAMAGE.
'
' ---------------------------------------------------------------------------
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owrss.asp,v $
'    $Revision: 1.2 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

Function RetrieveRSSFeed(pURL)
    Dim vXmlDoc, vRoot, vXslFilename

    On Error Resume Next
    'Response.Write("Retrieving " & pURL & "<br />")

    Set vXmlDoc = RetrieveXML(pURL)

    Set vRoot = vXmlDoc.documentElement

    ' determine the type of the feed
    If vRoot.NodeName = "rss" Then
        vXslFilename = "owrss091.xsl"
    Elseif vRoot.NodeName = "scriptingNews" Then
        vXslFilename = "owscriptingnews.xsl"
    Elseif vRoot.getAttribute("xmlns") = "http://my.netscape.com/rdf/simple/0.9/" Then
        vXslFilename = "owrss09.xsl"
    Elseif vRoot.getAttribute("xmlns") = "http://purl.org/rss/1.0/" Then
        ' TODO: find workaround for bug in MSXML v4
        If Not vRoot.selectSingleNode("item/ag:source") Is Nothing Then
            vXslFilename = "owrss10aggr.xsl"
        Else
            vXslFilename = "owrss10.xsl"
        End If
    Else
        Exit Function
    End If

    If cAllowAggregations Then
        Call gNamespace.Aggregate(pURL, vXmlDoc)
    End If

    RetrieveRSSFeed = gTransformer.TransformXmlDoc(vXmlDoc, vXslFilename)

    ' strip away any <script> elements, rigorously
    ' avoid running security risk of malicious javascript code
    RetrieveRSSFeed = s(RetrieveRSSFeed, "<script(.*?)script>", "", True, True)
End Function


' retrieve the XML data from the given URL
Function RetrieveXML(pURL)
    Dim vXmlDoc, vXmlHttp, vXmlStr, vPos, vPosEnd

    If MSXML_VERSION = 4 Then
        Set vXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
    Else
        Set vXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
    End If
    vXmlHttp.Open "GET", pURL, False
    vXmlHttp.send ""

    Set vXmlDoc = vXmlHttp.ResponseXML
    If vXmlDoc.xml = "" Then
        ' sometimes (quite often actually) an RSS feed can't be
        ' loaded into the DOM directly. This is usually because the
        ' feed is send with content-type text/plain instead of text/xml.
        ' For example, the RSS feeds from kuro5hin and salon.com won't
        ' load properly, resulting in an empty XML document object.
        '
        ' therefore, alternative method: first get the document as a string.
        vXmlStr = vXmlHttp.ResponseText

        ' unbelievable, but true, valid ISO-8859-1 characters in the vXmlStr
        ' variable won't load in a DOM document, here's an (imperfect) trick:
        vXmlStr = Server.HTMLEncode(vXmlStr)
        vXmlStr = Replace(vXmlStr, "&gt;", ">")
        vXmlStr = Replace(vXmlStr, "&lt;", "<")
        vXmlStr = Replace(vXmlStr, "&amp;", "&")
        vXmlStr = Replace(vXmlStr, "&quot;", """")
        vXmlStr = Replace(vXmlStr, "&#65535;", "?")

        ' the next stumbling block is that some contain the
        ' <!DOCTYPE ...> string which, although it's perfectly valid
        ' in XML world, for some really maddening reason won't load
        ' into an XML document object as well.
        '
        ' therefore, first strip it away
        vPos = InStr(vXmlStr, "<!DOCTYPE ")
        If vPos > 0 Then
            vPosEnd = InStr(vPos, vXmlStr, ">")
            If vPosEnd > 0 Then
                ' note: conveniently assume UTF-8 encoding
                vXmlStr = "<?xml version='1.0'?>" & Mid(vXmlStr, vPosEnd + 1)
            End If
        End If
        'Response.Write("<b><a href='" & pURL & "' target='_blank'>" & pURL & "</a></b><br />" & Server.HTMLEncode(vXmlStr) & "<br /><br />")

        ' and finally we can, hopefully, get it loaded as an xml document object
        If MSXML_VERSION = 4 Then
            Set vXmlDoc = Server.CreateObject("Msxml2.FreeThreadedDOMDocument.4.0")
        Else
            Set vXmlDoc = Server.CreateObject("Msxml2.FreeThreadedDOMDocument")
        End If
        vXmlDoc.async = False
        If Not vXmlDoc.loadXML(vXmlStr) Then
            ' sometimes this fails because of character endoding issues.
            ' if anyone knows a solid way to load XML feeds from other
            ' servers, plz let us know! -- LP
            'Response.Write("<p><b>Error</b> " & vXmlDoc.parseError.reason & " line: " & vXmlDoc.parseError.Line & " col: " & vXmlDoc.parseError.linepos & "</p>")
            Exit Function
        End If
    End If
    Set RetrieveXML = vXmlDoc
End Function


Function GetAggregation(pPage)
    Dim vXmlStr, vXmlDoc

    On Error Resume Next

    If Not IsObject(gAggregateURLs) Then
        Exit Function
    End If
    If gAggregateURLs.Count = 0 Then
        Exit Function
    End If

    vXmlStr = gNamespace.GetAggregation(gAggregateURLs)

    If MSXML_VERSION = 4 Then
        Set vXmlDoc = Server.CreateObject("Msxml2.FreeThreadedDOMDocument.4.0")
    Else
        Set vXmlDoc = Server.CreateObject("Msxml2.FreeThreadedDOMDocument")
    End If
    vXmlDoc.async = False
    If Not vXmlDoc.loadXML(vXmlStr) Then
        'Response.Write("<p><b>Error</b> " & vXmlDoc.parseError.reason & " line: " & vXmlDoc.parseError.Line & " col: " & vXmlDoc.parseError.linepos & "</p>")
        Exit Function
    End If

    vXmlStr = gTransformer.TransformXmlDoc(vXmlDoc, "owrss10aggr.xsl")

    ' strip away any <script> elements, rigorously
    ' avoid running security risk of malicious javascript code
    vXmlStr = s(vXmlStr, "<script(.*?)script>", "", True, True)

    GetAggregation = "<ow:aggregation href='" & CDATAEncode(gScriptName & "?p=" & pPage & "&a=rss") & "' " _
                   & "refreshURL='" & CDATAEncode(gScriptName & "?p=" & pPage & "&a=refresh&redirect=" & gPage) & "' "
    If Not vXmlDoc.documentElement.selectSingleNode("item/ag:timestamp") Is Nothing Then
        GetAggregation = GetAggregation & "last='" & vXmlDoc.documentElement.selectSingleNode("item/ag:timestamp").text & "' "
    End If
    If Request("refresh") = "" Then
        GetAggregation = GetAggregation & "fresh='false'"
    Else
        GetAggregation = GetAggregation & "fresh='true'"
    End If
    GetAggregation = GetAggregation & ">" & vXmlStr & "</ow:aggregation>"
End Function
%>