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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owwikify.asp,v $
'    $Revision: 1.5 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

'______________________________________________________________________________________________________________
Function Wikify(pText)
    Dim vText
    vText = pText

    gIncludingAsTemplate = False
    If gIncludeLevel = 0 Then
        Set gRaw            = New Vector
        Set gBracketIndices = New Vector
        Set gTOC            = New TableOfContents

        If gAction <> "edit" And Not cEmbeddedMode Then
            If Left(vText, 1) = "#" Then
                If m(vText, "^#RANDOMPAGE", False, False) Then
                    ActionRandomPage()
                Elseif m(vText, "^#REDIRECT\s+", False, False) And Request("redirect") = "" Then
                    gTemp = InStr(10, vText, vbCR)
                    If gTemp > 0 Then
                        gTemp = Trim(Mid(vText, 10, gTemp - 10))
                    Else
                        gTemp = Trim(Mid(vText, 10))
                    End If
                    Response.Redirect(gScriptName & "?a=" & gAction & "&p=" & Server.URLEncode(gTemp) & "&redirect=" & Server.URLEncode(FreeToNormal(gPage)))
                Elseif m(vText, "^#INCLUDE_AS_TEMPLATE", False, False) Then
                    vText = Mid(vText, 21)
                Elseif m(vText, "^#MINOREDIT", False, False) Then
                    vText = Mid(vText, 11)
                Elseif m(vText, "^#DEPRECATED", False, False) Then
                    StoreRaw("#DEPRECATED")
                    vText = sReturn & Mid(vText, 12)
                End If
                vText = MyWikifyProcessingInstructions(vText)
            End If
        End If
    Else
        If gAction <> "edit" And Not cEmbeddedMode Then
            If Left(vText, 1) = "#" Then
                If m(vText, "^#INCLUDE_AS_TEMPLATE", False, False) Then
                    vText = Mid(vText, 21)
                    gIncludingAsTemplate = True
                End If
            End If
        End If
    End If

    vText = MultiLineMarkup(vText)  ' Multi-line markup
    vText = WikiLinesToHtml(vText)  ' Line-oriented markup

    vText = s(vText, gFS & "(\d+)" & gFS, "&GetRaw($1)", False, True)  ' Restore saved text
    vText = s(vText, gFS & "(\d+)" & gFS, "&GetRaw($1)", False, True)  ' Restore nested saved text

    If gIncludeLevel = 0 Then
        If cUseHeadings Then
            vText = s(vText, gFS & "(\=+)[ \t]+(.*?)[ \t]+\=+ " & gFS, "&GetWikiHeading($1, $2)", False, True)
            vText = Replace(vText, gFS & "TOC" & gFS, gTOC.GetTOC)
        End If

        If InStr(gMacros, "Footnote") > 0 Then
            vText = InsertFootnotes(vText)
        End If

        vText = MyLastMinuteChanges(vText)
        Set gRaw            = Nothing
        Set gBracketIndices = Nothing
        Set gTOC            = Nothing
    End If

    Wikify = vText
End Function


'______________________________________________________________________________________________________________
Function MultiLineMarkup(pText)
    pText = Replace(pText & "", Chr(9), Space(8))
    'pText = Replace(pText, gFS, "")    ' remove separators

    If cRawHtml Then
        pText = s(pText, "<html>([\s\S]*?)<\/html>", "&StoreHtml($1)", True, True)
    End If
    If cMathML Then
        pText = s(pText, "<math>([\s\S]*?)<\/math>", "&StoreMathML($1)", True, True)
    End If

    pText = MyMultiLineMarkupStart(pText)

    pText = QuoteXml(pText)
    If cRawHtml Then
        ' transform our field separator back
        pText = Replace(pText, "&#179;", gFS)
    End If
    pText = s(pText, " \\ *\r?\n", "", False, True)  ' Join lines with backslash at end

    ' The <nowiki> tag stores text with no markup (except quoting HTML)
    pText = s(pText, "\&lt;nowiki\&gt;([\s\S]*?)\&lt;\/nowiki\&gt;", "&StoreRaw($1)", True, True)

    ' <code></code> and {{{ }}} do the same thing.
    pText = s(pText, "\{\{\{(.*?)\}\}\}", "&StoreRaw(""<tt>"" & $1 & ""</tt>"")", True, True)
    pText = s(pText, "\&lt;code\&gt;(.*?)\&lt;\/code\&gt;", "&StoreRaw(""<tt>"" & $1 & ""</tt>"")", True, True)
    pText = s(pText, "\{\{\{([\s\S]*?)\}\}\}", "&StoreCode($1)", True, True)
    pText = s(pText, "\&lt;code\&gt;([\s\S]*?)\&lt;\/code\&gt;", "&StoreCode($1)", True, True)
    pText = s(pText, "\&lt;pre\&gt;([\s\S]*?)\&lt;\/pre\&gt;", "<pre>$1</pre>", True, True)

    If cHtmlTags Then
        ' Scripting is currently possible with these tags, so they are *not* particularly "safe".
        Dim vTag
        For Each vTag In Split("b,i,u,font,big,small,sub,sup,h1,h2,h3,h4,h5,h6,cite,code,em,s,strike,strong,tt,var,div,center,blockquote,ol,ul,dl,table,caption,br,p,hr,li,dt,dd,tr,td,th", ",")
            pText = s(pText, "\&lt;" & vTag & "(\s[^<>]+?)?\&gt;([\s\S]*?)\&lt;\/" & vTag & "\&gt;", "<" & vTag & "$1>$2</" & vTag & ">", True, True)
        Next
        For Each vTag In Split("br,p,hr,li,dt,dd,tr,td,th", ",")
            pText = s(pText, "\&lt;" & vTag & "(\s[^<>]+?)?\&gt;", "<" & vTag & "$1 />", True, True)
        Next
    End If

    If cHtmlLinks Then
        pText = s(pText, "\&lt;a\s([^<>]+?)\&gt;([\s\S]*?)\&lt;\/a\&gt;", "&StoreHref($1, $2)", True, True)
    End If

    If IsObject(gAggregateURLs) Then
        ' we are in the process of refreshing RSS feeds
        If m(gMacros, "Include", True, True) Then
            pText = s(pText, "\&lt;(Include)(\(.*?\))?(?:\s*\/)?\&gt;", "&ExecMacro($1, $2)", True, True)
        End If
        pText = s(pText, "\&lt;(Syndicate)(\(.*?\))?(?:\s*\/)?\&gt;", "&ExecMacro($1, $2)", True, True)
        MultiLineMarkup = pText
        Exit Function
    End If

    ' process macro's
    pText = s(pText, "\&lt;(" & gMacros & ")(\(.*?\))?(?:\s*\/)?\&gt;", "&ExecMacro($1, $2)", True, True)

    If cFreeLinks Then
        pText = s(pText, "\[\[" & gFreeLinkPattern & "(?:\|([^\]]+))*\]\]", "&StoreFreeLink($1, $2)", False, True)
    End If

    ' Links like [URL] and [URL text of link]
    pText = s(pText, "\[" & gUrlPattern & "(\s+[^\]]+)*\]", "&StoreBracketUrl($1, $2)", False, True)
    pText = s(pText, "\[" & gInterLinkPattern & "(\s+[^\]]+)*\]", "&StoreInterPage($1, $2, True)", False, True)
    pText = s(pText, "\[" & gISBNPattern & "([^\]]+)*\]", "&StoreISBN($1, $2, True)", False, True)

    If cAllowAttachments Then
        Dim vAttachmentPattern
        If IsObject(gCurrentWorkingPages) Then
            ' we're including a page
            Set gTemp = gNamespace.GetPageAndAttachments(gCurrentWorkingPages.Top(), 0, True, False)
        Else
            Set gTemp = gNamespace.GetPageAndAttachments(gPage, gRevision, True, False)
        End If
        vAttachmentPattern = gTemp.GetAttachmentPattern()
        If vAttachmentPattern <> "" Then
            pText = s(pText, "\[(" & gTemp.GetAttachmentPattern & ")(\s+[^\]]+)*\]", "&StoreBracketAttachmentLink($1, $2)", False, True)
        End If
    End If

    If cWikiLinks And cBracketText And cBracketWiki Then
        ' Local bracket-links
        pText = s(pText, "\[" & "(#?)" & gLinkPattern & "(\s+[^\]]+?)\]", "&StoreBracketWikiLink($1, $2, $3)", False, True)
    End If

    pText = s(pText, gUrlPattern, "&StoreUrl($1)", False, True)
    pText = s(pText, gInterLinkPattern, "&StoreInterPage($1, """", False)", False, True)
    pText = s(pText, gMailPattern, "&StoreMail($1)", False, True)
    pText = s(pText, gISBNPattern, "&StoreISBN($1, """", False)", False, True)

    If cAllowAttachments Then
        If IsObject(gCurrentWorkingPages) Then
            ' we're including a page
            Set gTemp = gNamespace.GetPageAndAttachments(gCurrentWorkingPages.Top(), 0, True, False)
        Else
            Set gTemp = gNamespace.GetPageAndAttachments(gPage, gRevision, True, False)
        End If
        vAttachmentPattern = gTemp.GetAttachmentPattern()
        If vAttachmentPattern <> "" Then
            pText = s(pText, "(" & gTemp.GetAttachmentPattern & ")", "&StoreAttachmentLink($1)", False, True)
        End If
    End If

    pText = s(pText, "-{4,}", "<hr size=""1""/>", False, True)
    pText = s(pText, "\&gt;\&gt;([\s\S]*?)\&lt;\&lt;", "<center>$1</center>", False, True)

    If cNewSkool Then
        pText = s(pText, "\*\*([^\s\*].*?)\*\*", "<b>$1</b>", False, True)
        pText = s(pText, "\/\/([^\s\/].*?)\/\/", "<i>$1</i>", False, True)
        pText = s(pText, "__([^\s_].*?)__", "<u>$1</u>", False, True)
        pText = s(pText, "--([^\s-].*?)--", "<strike>$1</strike>", False, True)
        pText = s(pText, "!!([^\s!].*?)!!", "<big>$1</big>", False, True)
        pText = s(pText, "\^\^([^\s\^].*?)\^\^", "<sup>$1</sup>", False, True)
        pText = s(pText, "vv([^\sv].*?)vv", "<sub>$1</sub>", False, True)
        'pText = s(pText, " --", " &#173;", False, True)
    End If

    If cUseHeadings And cWikifyHeaders = 0 Then
        pText = s(pText, gHeaderPattern, "&StoreWikiHeading($1, $2, $3)", False, True)
    End If

    If cWikiLinks Then
        If OPENWIKI_STOPWORDS <> "" Then
            gStopWords = gNamespace.GetPage(OPENWIKI_STOPWORDS, 0, True, False).Text
            gStopWords = Replace(gStopWords & "", Chr(9), " ")
            gStopWords = Replace(gStopWords, gFS, "")    ' remove separators
            gStopWords = Replace(gStopWords, vbCR, " ")
            gStopWords = Replace(gStopWords, vbLF, " ")
            gStopWords = Trim(gStopWords)
            gStopWords = s(gStopWords, "\s+", "|", False, True)
        End If

        If gStopWords <> "" Then
            pText = s(pText, "\b(" & gStopWords & ")\b", "&StoreRaw($1)", True, True)
        End If

        If cNewSkool Then
            pText = s(pText, "(~?)" & gLinkPattern, "&GetWikiLink($1, $2, """")", False, True)
        Else
            pText = s(pText, gLinkPattern, "&GetWikiLink("""", $1, """")", False, True)
        End If
    End If

    If cOldSkool Then
        ' The quote markup patterns avoid overlapping tags (with 5 quotes)
        ' by matching the inner quotes for the strong pattern.
        pText = Replace(pText, "''''''", "")
        pText = s(pText, "('*)'''(.*?)'''", "$1<strong>$2</strong>", False, True)
        pText = s(pText, "''(.*?)''", "<em>$1</em>", False, True)
    End If

    If Not cHtmlTags Then
        ' I disabled this because I don't like this way of quoting
        ' Enabling this forces editors to use "correct" HTML, i.e. XHTML.
        ' E.g. <b><i>bla</b></i> will fail, because it's not valid XHTML. -- LawrencePit
        'pText = s(pText, "\&lt;b\&gt;(.*?)\&lt;\/b\&gt;", "<b>$1</b>", True, True)
        'pText = s(pText, "\&lt;i\&gt;(.*?)\&lt;\/i\&gt;", "<i>$1</i>", True, True)
        'pText = s(pText, "\&lt;u\&gt;(.*?)\&lt;\/u\&gt;", "<u>$1</u>", True, True)
        'pText = s(pText, "\&lt;strong\&gt;(.*?)\&lt;\/strong\&gt;", "<strong>$1</strong>", True, True)
        'pText = s(pText, "\&lt;em\&gt;(.*?)\&lt;\/em\&gt;", "<em>$1</em>", True, True)
    End If

    If cEmoticons Then
        pText = s(pText, "\s\:\-?\)($|\s)", " <img src=""" & OPENWIKI_ICONPATH & "/emoticon-smile.gif"" width=""14"" height=""12"" alt=""""/>$1", True, True)
        pText = s(pText, "\s\;\-?\)($|\s)", " <img src=""" & OPENWIKI_ICONPATH & "/emoticon-wink.gif"" width=""14"" height=""12"" alt=""""/>$1", True, True)
        pText = s(pText, "\s\:\-?\(($|\s)", " <img src=""" & OPENWIKI_ICONPATH & "/emoticon-sad.gif"" width=""14"" height=""12"" alt=""""/>$1", True, True)
        pText = s(pText, "\s\:\-?\|($|\s)", " <img src=""" & OPENWIKI_ICONPATH & "/emoticon-ambivalent.gif"" width=""14"" height=""12"" alt=""""/>$1", True, True)
        pText = s(pText, "\s\:\-?D($|\s)", " <img src=""" & OPENWIKI_ICONPATH & "/emoticon-laugh.gif"" width=""14"" height=""12"" alt=""""/>$1", True, True)
        pText = s(pText, "\s\:\-?O($|\s)", " <img src=""" & OPENWIKI_ICONPATH & "/emoticon-surprised.gif"" width=""14"" height=""12"" alt=""""/>$1", True, True)
        pText = s(pText, "\s\:\-?P($|\s)", " <img src=""" & OPENWIKI_ICONPATH & "/emoticon-tongue-in-cheek.gif"" width=""14"" height=""12"" alt=""""/>$1", True, True)
        pText = s(pText, "\s\:\-?S($|\s)", " <img src=""" & OPENWIKI_ICONPATH & "/emoticon-unsure.gif"" width=""14"" height=""12"" alt=""""/>$1", True, True)
        pText = s(pText, "(^|\s)\(([Y|N|L|U|K|G|F|P|B|D|T|C|I|H|S|8|E|M])\)($|\s)", "$1<img src=""" & OPENWIKI_ICONPATH & "/emoticon-$2.gif"" width=""14"" height=""12"" alt=""""/>$3", True, True)
        pText = s(pText, "(^|\s)\(\*\)($|\s)", "$1<img src=""" & OPENWIKI_ICONPATH & "/emoticon-star.gif"" width=""14"" height=""12"" alt=""""/>$2", True, True)
        pText = s(pText, "(^|\s)\(\@\)($|\s)", "$1<img src=""" & OPENWIKI_ICONPATH & "/emoticon-cat.gif"" width=""14"" height=""12"" alt=""""/>$2", True, True)
        pText = s(pText, "(^|\s)\/i\\($|\s)", "$1<img src=""" & OPENWIKI_ICONPATH & "/icon-info.gif"" width=""16"" height=""16"" alt=""""/>$2", True, True)
        pText = s(pText, "(^|\s)\/w\\($|\s)", "$1<img src=""" & OPENWIKI_ICONPATH & "/icon-warning.gif"" width=""16"" height=""16"" alt=""""/>$2", True, True)
        pText = s(pText, "(^|\s)\/s\\($|\s)", "$1<img src=""" & OPENWIKI_ICONPATH & "/icon-error.gif"" width=""16"" height=""16"" alt=""""/>$2", True, True)
    End If

    If cUseHeadings And cWikifyHeaders Then
        pText = s(pText, gHeaderPattern, "&StoreWikiHeading($1, $2, $3)", False, True)
    End If

    pText = MyMultiLineMarkupEnd(pText)

    MultiLineMarkup = pText
End Function


'______________________________________________________________________________________________________________
Function WikiLinesToHtml(pText)
    Dim vTagStack, vRegEx, vMatch, vMatches, vLine
    Dim vFirstChar, vCode, vDepth, vPos, vStart, vAttrs
    Dim vInTable
    Dim vText

    vText = ""
    vDepth = 0
    vInTable = 0
    Set vTagStack = new Vector
    Set vRegEx = New RegExp
    vRegEx.IgnoreCase = False
    vRegEx.Global = True
    vRegEx.Pattern = ".+"
    Set vMatches = vRegEx.Execute(pText)
    For Each vMatch In vMatches
        'vLine = vMatch.Value
        vLine = RTrim(Replace(vMatch.Value, vbCR, ""))
        vLine = s(vLine, "^\s*$", "<p></p>", False, True)  ' Blank lines

        ' The following piece of code is not as bad as you could hope for
        vFirstChar = Left(vLine, 1)
        If (vFirstChar = " ") Or (vFirstChar = Chr(8)) Then

            If (vDepth = 0) And (vInTable > 0) Then
                vText = vText & vbCRLF & "</table>" & vbCRLF
                vInTable = 0
            End If

            vAttrs = ""
            gListSet = False
            vLine = s(vLine, "^(\s+)\;(.*?) \:", "&SetListValues(True, $1, ""<dt>"" & $2 & ""</dt><dd>"")", False, True)
            If gListSet Then
                vLine = vLine & "</dd>"
                vCode = "dl"
                vDepth = Len(gDepth) / 2
            Else
                vLine = s(vLine , "^(\s+)\:\s(.*?)$", "&SetListValues(True, $1, ""<dt> </dt><dd>"" & $2 & ""</dd>"")", False, True)
                If gListSet Then
                  vCode = "dl"
                  vDepth = Len(gDepth) / 2
                Else
                    vLine = s(vLine, "^(\s+)\*\s(.*?)$", "&SetListValues(True, $1, ""<li>"" & $2 & ""</li>"")", False, True)
                    If gListSet Then
                        vCode  = "ul"
                        vDepth = Len(gDepth) / 2
                    Else
                        vLine = s(vLine, "^(\s+)([0-9aAiI]\.(?:#\d+)? )", "&SetListValues(True, $1, $2)", False, True)
                        If gListSet Then
                            vPos   = InStr(vLine, " ")
                            vCode  = Left(vLine, vPos - 1)
                            vLine  = "<li>" & Mid(vLine, vPos + 1) & "</li>"

                            vPos   = InStr(vCode, "#")
                            vStart = ""
                            If vPos > 0 Then
                                vStart = "start=""" & Mid(vCode, vPos + 1) & """"
                            End If
                            vCode = Left(vCode, 1)
                            If IsNumeric(vCode) Then
                                vAttrs = " type=""1"""
                            Else
                                vAttrs = " type=""" & vCode & """"
                            End If
                            If vStart <> "" Then
                                vAttrs = vAttrs & " " & vStart
                            End If
                            vCode  = "ol"
                            vDepth = Len(gDepth) / 2
                        Elseif vDepth > 0 And vCode <> "pre" Then
                            Dim vTemp
                            vTemp = Trim(vLine)
                            If (Left(vTemp, 2) = "||") And (Right(vTemp, 2) = "||") Then
                                vLine = vTemp
                            Elseif vInTable = 0 Then
                                vText = vText & "<br />"
                            End If
                        Else
                            vCode = "pre"
                            vDepth = 1
                        End If
                    End If
                End If
            End If
        Else
            If (vDepth > 0) And (vInTable > 0) Then
                vText = vText & vbCRLF & "</table>" & vbCRLF
                vInTable = 0
            End If

            vDepth = 0
        End If

        Do While (vTagStack.Count > vDepth)
            vText = vText & "</" & vTagStack.Pop() & ">" & vbCRLF
        Loop

        If (vDepth > 0) Then
            If vDepth > gIndentLimit Then
                vDepth = gIndentLimit
            End If
            Do While (vTagStack.Count < vDepth)
                vTagStack.Push(vCode)
                vText = vText & "<" & vCode & vAttrs & ">" & vbCRLF
            Loop
            If Not vTagStack.IsEmpty Then
                If vTagStack.Top <> vCode Then
                    vText = vText & "</" & vTagStack.Pop() & ">"  & vbCRLF & "<" & vCode & vAttrs & ">"
                    vTagStack.Push(vCode)
                End If
            End If
        End If

        If Left(vLine, 2) = "||" And Right(vLine, 2) = "||" Then
            ' tables
            Dim vTR, vTD, vColSpan, vNrOfTDs, vResult, vSaveReturn
            vTR = vLine
            vNrOfTDs = 0
            vResult = ""

            Do While vTR <> ""
                gListSet = False
                vTD = s(vTR, "^(\|{2,})(.*?)\|\|", "&SetListValues(True, $1, $2)", False, True)
                If gListSet Then
                    vColSpan = Int(Len(gDepth) / 2)
                    vNrOfTDs = vNrOfTDs + vColSpan
                    If vColSpan = 1 Then
                        vColSpan = "<td class=""wiki"">"
                    Else
                        vColSpan = "<td class=""wiki"" align=""center"" colspan=""" & vColSpan & """>"
                    End If
                    vSaveReturn = sReturn
                    If Trim(sReturn) = "" Then
                        sReturn = "&#160;"
                    End If
                    vResult = vResult & vColSpan & sReturn & "</td>"
                    'Response.Write("GOT: " & Server.HTMLEncode(vResult) & "<br>")
                    vTR = Mid(vTR, Len(gDepth) + Len(vSaveReturn) + 1)
                Else
                    vTR = ""
                End If
            Loop

            If (vInTable > 0) And (vInTable <> vNrOfTDs) Then
                vText = vText & vbCRLF & "</table>" & vbCRLF
                vInTable = 0
            End If

            If vInTable = 0 Then
                vText = vText & "<table cellspacing=""0"" cellpadding=""2"" border=""1"" class=""wiki"">"
                vInTable = vNrOfTDs
            End If
            vText = vText & vbCRLF & "<tr class=""wiki"">" & vResult & "</tr>"
        Elseif vInTable > 0 Then
            vText = vText & vbCRLF & "</table>" & vbCRLF
            vInTable = 0
        End If

        If vInTable = 0 Then
            vText = vText & vLine & vbCRLF
        End If
    Next

    If vInTable > 0 Then
        vText = vText & vbCRLF & "</table>" & vbCRLF
    End If

    Do While Not vTagStack.IsEmpty
        vText = vText & "</" & vTagStack.Pop() & ">" & vbCRLF
    Loop

    Set vRegEx = Nothing
    Set vTagStack = Nothing

    WikiLinesToHtml = vText
End Function

Dim gListSet, gDepth
Sub SetListValues(pListSet, pDepth, pText)
    gListSet = pListSet
    gDepth   = pDepth
    sReturn  = pText
End Sub

'______________________________________________________________________________________________________________
Function QuoteXml(ByRef pText)
    QuoteXml = Replace(pText, "&", "&amp;")
    QuoteXml = Replace(QuoteXml, "<", "&lt;")
    QuoteXml = Replace(QuoteXml, ">", "&gt;")

    ' In XML data HTML character references are invalid (unless these are
    ' defined in the DTD). Special characters can be entered in XML without
    ' the use of character references. Make sure you've set the constant
    ' OPENWIKI_ENCODING correct though in owconfig.asp and also the encoding
    ' attribute at the first line of the stylesheets.
    If cAllowCharRefs Then
        QuoteXml = s(QuoteXml, "\&amp;([#a-zA-Z0-9]+);", "&StoreCharRef($1)", False, True)
    End If
End Function

Function CDATAEncode(pText)
    If pText <> "" Then
        CDATAEncode = Replace(pText, "&", "&amp;")
        CDATAEncode = Replace(CDATAEncode, "<", "&lt;")
        CDATAEncode = Replace(CDATAEncode, "'", "&apos;")
    End If
End Function

Function PCDATAEncode(pText)
    If pText <> "" Then
        PCDATAEncode = Replace(pText, "&", "&amp;")
        PCDATAEncode = Replace(PCDATAEncode, "<", "&lt;")
        PCDATAEncode = Replace(PCDATAEncode, "]]>", "]]&gt;")
    End If
End Function

Function URLDecode(pURL)
    Dim vPos
    If pURL <> "" Then
        pURL = Replace(pURL, "+", " ")
        vPos = InStr(pURL, "%")
        Do While vPos > 0
            pURL = Left(pURL, vPos - 1) _
                 & Chr(CLng("&H" & Mid(pURL, vPos + 1, 2))) _
                 & Mid(pURL, vPos + 3)
            vPos = InStr(vPos + 1, pURL, "%")
        Loop
    End If
    URLDecode = pURL
End Function

'______________________________________________________________________________________________________________
Sub StoreRaw(pText)
    gRaw.Push(pText)
    sReturn = gFS & (gRaw.Count - 1) & gFS
End Sub

Sub GetRaw(pIndex)
    sReturn = gRaw.ElementAt(pIndex)
End Sub

Sub StoreCharRef(pText)
    StoreHtml("&" & pText & ";")
End Sub

Sub StoreHtml(pText)
    StoreRaw("<ow:html><![CDATA[" & Replace(pText, "]]>", "]]&gt;") & "]]></ow:html>")
End Sub

Sub StoreMathML(pText)
    StoreRaw("<ow:math><![CDATA[" & Replace(pText, "]]>", "]]&gt;") & "]]></ow:math>")
End Sub

Sub StoreCode(pText)
    StoreRaw("<pre class=""code"">" & s(pText, "'''(.*?)'''", "<b>$1</b>", False, True) & "</pre>")
End Sub

Sub StoreMail(pText)
    StoreRaw("<a href=""mailto:" & pText & """ class=""external"">" & pText & "</a>")
End Sub

Sub StoreUrl(pURL)
    Call UrlLink(pURL)
    StoreRaw(gTempLink)
    sReturn = sReturn & gTempJunk
End Sub

Sub StoreBracketUrl(pURL, pText)
    If pText = "" Then
        If cUseLinkIcons Then
            pText = pURL
        End If
    Else
        If cBracketText = 0 Then
            sReturn = "[" & pURL & " " & pText & "]"
            Exit Sub
        End If
    End If
    StoreRaw(GetExternalLink(pURL, pText, "", True))
End Sub

Sub StoreHref(pAnchor, pText)
    Dim vLink
    vLink = "<a " & pAnchor
    If cExternalOut Then
        If Not m(pAnchor, " target=\""", True, True) Then
            vLink = vLink & " target=""_blank"""
        End If
    End If
    If Not m(pAnchor, " class=\""", True, True) Then
        vLink = vLink & " class=""external"""
    End If
    vLink = vLink & ">"
    If cUseLinkIcons Then
        vLink = vLink & "<img src=""" & OPENWIKI_ICONPATH & "/wiki-http.gif"" border=""0"" hspace=""4"" alt=""""/>"
    End If
    vLink = vLink & pText & "</a>"
    StoreRaw(vLink)
End Sub

Sub StoreFreeLink(pID, pText)
    ' trim spaces before/after subpages
    pID = s(pID, "\s*\/\s*", "/", False, True)
    gTemp = GetWikiLink("", Trim(pID), Trim(pText))
    If Left(gTemp, 1) <> "<" Then
        sReturn = "[[" & pID & pText & "]]"
    Else
        StoreRaw(gTemp)
    End If
End Sub

Sub StoreBracketWikiLink(pPrefix, pID, pText)
    If pID = gPage Then
        ' don't link to oneself
        sReturn = pText
    Else
        gTemp = GetWikiLink(pPrefix, pID, LTrim(pText))
        If Left(gTemp, 1) <> "<" Then
            sReturn = "[" & pPrefix & pID & pText & "]"
        Else
            StoreRaw(gTemp)
        End If
    End If
End Sub

Sub StoreInterPage(pID, pText, pUseBrackets)
    Dim vPos, vSite, vRemotePage, vURL, vTemp
    If pUseBrackets Then
        gTempLink = pID
        gTempJunk = ""
    Else
        SplitUrlPunct(pID)
    End If
    vPos = InStr(gTempLink, ":")
    If vPos > 0 Then
        vSite = Left(gTempLink, vPos - 1)
        vRemotePage = Mid(gTempLink, vPos + 1)
        vURL = gNamespace.GetInterWiki(vSite)
    End If
    If vURL = "" Then
        sReturn = pID & pText
        If pUseBrackets Then
            sReturn = "[" & sReturn & "]"
        End If
    Else
        If pText = "" Then
            If pUseBrackets And cBracketIndex And (cUseLinkIcons = 0) Then
                pText = ""
            Else
                ' pText = Mid(pID, Len(vSite) + 2)
                ' pText = pID
                pText = gTempLink
            End If
        Elseif cBracketText = 0 Then
            If pUseBrackets Then
                sReturn = "[" & pID & pText & "]"
                Exit Sub
            End If
        End If
        If vPos > 0 Then
            If InStr(vURL, "$1") > 0 Then
                vURL = Replace(vURL, "$1", vRemotePage)
            Else
                vURL = vURL & vRemotePage
            End If
        Else
            vURL = vURL & vRemotePage
        End If
        vURL = Replace(vURL, "&", "&amp;")
        vURL = Replace(vURL, "&amp;amp;", "&amp;")  ' correction back
        If vSite = "This" Then
            StoreRaw("<ow:link name='" & pText & "' href='" & vURL & "' date='" & FormatDateISO8601(Now()) & "'>" & pText & "</ow:link>" & gTempJunk)
        Else
            StoreRaw(GetExternalLink(vURL, pText, vSite, pUseBrackets) & gTempJunk)
        End If
    End If
End Sub

Sub StoreISBN(pNumber, pText, pUseBrackets)
    If pText <> "" And cBracketText = 0 And pUseBrackets Then
        sReturn = "[ISBN" & pNumber & pText & "]"
    Else
        Dim vRawPrint, vNumber, vText
        vRawPrint = Replace(pNumber, " ", "")
        vNumber = Replace(vRawPrint, "-", "")

        If Len(vNumber) = 11 Then
            If UCase(Right(vNumber, 1)) = "X" Then
                pText = Right(vNumber, 1) & pText
                vNumber = Left(vNumber, 10)
            End If
        End If

        If Len(vNumber) <> 10 Then
            If pText = "" Then
                sReturn = "ISBN " & pNumber
            Else
                sReturn = "[ISBN " & pNumber & pText & "]"
            End If
        Else
            If pText = "" Then
                If pUseBrackets And cBracketIndex And (cUseLinkIcons = 0) Then
                    vText = ""
                Else
                    vText = "ISBN " & vRawPrint
                End If
            Else
                vText = pText
            End If
            sReturn = GetExternalLink("http://www.amazon.com/exec/obidos/ISBN=" & vNumber, vText, "Amazon", pUseBrackets) _
                    & " (" & GetExternalLink("http://shop.barnesandnoble.com/bookSearch/isbnInquiry.asp?isbn=" & vNumber, "alternate", "Barnes & Noble", False) _
                    & ", " & GetExternalLink("http://www1.fatbrain.com/asp/bookinfo/bookinfo.asp?theisbn=" & vNumber, "alternate", "FatBrain", False) & ")"

            If (pText = "") And (Right(pNumber, 1) = " ") Then
                sReturn = sReturn & " "
            End If
            StoreRaw(sReturn)
        End If
    End If
End Sub

Sub StoreWikiHeading(pSymbols, pText, pTrailer)
    StoreRaw(gFS & pSymbols & " " & pText & " " & pSymbols & " " & gFS)
    sReturn = sReturn & pTrailer
End Sub

Sub GetWikiHeading(pSymbols, pText)
    Dim vLevel, vTemp
    vLevel = Len(pSymbols)
    If vLevel > 6 Then
        vLevel = 6
    End If
    vTemp = s(pText, "<ow:link name='(.*?)' href=.*?</ow:link>", "$1", False, False)
    Call gTOC.AddTOC(vLevel, "<li><a href=""#h" & gTOC.Count & """>" & vTemp & "</a></li>")
    sReturn = "<a name=""h" & (gTOC.Count - 1) & """/><h" & vLevel & ">" & pText & "</h" & vLevel & ">"
End Sub


Sub StoreBracketAttachmentLink(pName, pText)
    gTemp = AttachmentLink(pName, pText)
    If gTemp = "" Then
        sReturn = "[" & pName & " " & pText & "]"
    Else
        StoreRaw(gTemp)
    End If
End Sub

Sub StoreAttachmentLink(pName)
    gTemp = AttachmentLink(pName, "")
    If gTemp = "" Then
        sReturn = pName
    Else
        StoreRaw(gTemp)
    End If
End Sub



'______________________________________________________________________________________________________________
Function PrettyWikiLink(pID)
    If cPrettyLinks Then
        PrettyWikiLink = s(pID, "([a-z\xdf-\xff0-9])([A-Z\xc0-\xde]+)", "$1 $2", False, True)
    Else
        PrettyWikiLink = pID
    End If
    If cFreeLinks Then
        PrettyWikiLink = Replace(PrettyWikiLink, "_", " ")
    End If
End Function


Function GetWikiLink(pPrefix, pID, pText)
    Dim vID, vPage, vAnchor, vTemplate, vTemp

    If pPrefix = "~" Then
        GetWikiLink = pID
        sReturn = GetWikiLink
        Exit Function
    End If

    If pPrefix = "#" Then
        vAnchor = "#" & pID
        pID = gPage
    Elseif pID = gPage Then
        ' don't link to oneself
        GetWikiLink = PrettyWikiLink(pID)
        sReturn = GetWikiLink
        Exit Function
    End If

    ' detect anchor
    vTemp = InStr(pID, "#")
    If vTemp > 0 Then
        vAnchor = Mid(pID, vTemp)
        pID = Left(pID, vTemp - 1)
    End If

    ' detect template
    vTemp = InStr(pID, "-&gt;")
    If vTemp > 0 then
        vTemplate = Left(pID, vTemp - 1)
        pID = Mid(pID, vTemp + 5)
    End if

    vID = AbsoluteName(pID)

    Set vPage = gNamespace.GetPage(vID, 0, False, False)
    vPage.Anchor = vAnchor
    If vPage.Exists Then
        If pText = "" Then
            GetWikiLink = vPage.ToLinkXML(PrettyWikiLink(pID), vTemplate, True)
        Else
            GetWikiLink = vPage.ToLinkXML(pText, vTemplate, False)
        End If
    Else
        If cReadOnly Or gAction = "print" Then
            GetWikiLink = pID & vAnchor
        Else
            If pText = "" Then
                pText = pID
            End If

            If cFreeLinks Then
                If InStr(pText, " ") > 0 Then
                    pText = "[" & pText & "]" ' Add brackets so boundaries are obvious
                End If
            End If

            ' non existent link
            GetWikiLink = vPage.ToLinkXML(pText, vTemplate, True)
        End If
    End If
    sReturn = GetWikiLink
End Function


Function AbsoluteName(pID)
    Dim vPos, vTemp, vCurrentPage, vMainpage

    If Not gIncludingAsTemplate And IsObject(gCurrentWorkingPages) Then
        vCurrentPage = gCurrentWorkingPages.Top()
    Else
        vCurrentPage = gPage
    End If

    ' asbolute subpage
    vPos = InStr(vCurrentPage, "/")
    If vPos > 0 Then
        vMainpage = Left(vCurrentPage, vPos - 1)
    Else
        vMainpage = vCurrentPage
    End If
    AbsoluteName = s(pID, "^/", vMainpage & "/", False, True)

    ' relative subpage
    AbsoluteName = s(AbsoluteName, "^\./", vCurrentPage & "/", False, True)

    If cFreeLinks Then
        AbsoluteName = FreeToNormal(AbsoluteName)
    End If
End Function


Function FreeToNormal(pID)
    Dim vID
    vID = Replace(pID, " ", "_")
    vID = UCase(Left(vID, 1)) & Mid(vID, 2)
    If InStr(vID, "_") > 0 Then
        vID = s(vID, "__+", "_",  False, True)
        vID = s(vID, "^_", "", False, True)
        vID = s(vID, "_$", "", False, True)
        If cUseSubpage Then
            vID = s(vID, "_\/", "/", False, True)
            vID = s(vID, "\/_", "/", False, True)
        End If
    End If
    If cFreeUpper Then
        vID = s(vID, "([-_\.,\(\)\/])([a-z])", "&Capitalize($1, $2)", False, True)
    End If
    FreeToNormal = vID
End Function

Sub Capitalize(pChars, pWord)
    sReturn = pChars & UCase(Left(pWord, 1)) & Mid(pWord, 2)
End Sub


Function GetExternalLink(pURL, pText, pTitle, pUseBrackets)
    Dim vLink, vLinkedImage, vTemp
    If pUseBrackets And pText = "" Then
        If cBracketIndex Then
            pText = "[" & GetBracketUrlIndex(pURL) & "]"
        Else
            pText = pURL
        End If
    Else
        pText = Trim(pText)
    End If

    If cAllowAttachments And (Left(pURL, 13) = "attachment://") Then
        If pUseBrackets And cShowBrackets Then
            pText = "[" & pText & "]"
        End If
        GetExternalLink = AttachmentLink(Mid(pURL, 14), pText)
        If GetExternalLink = "" Then
            GetExternalLink = "[" & pURL & " " & pText & "]"
        End If
        Exit Function
    End If

    vLink = "<a href='" & pURL & "' class='external'"
    If cExternalOut Then
        vLink = vLink & " target='_blank'"
    End If
    If pTitle <> "" Then
        vLink = vLink & " title='" & CDATAEncode(pTitle) & "'"
    End If
    vLink = vLink & ">"

    vLinkedImage = False
    If pText <> "" Then
        If m(pText, gImagePattern, False, True) Then
            pText = "<img src=""" & pText & """ border=""0"" alt=""""/>"
            vLinkedImage = True
        End If
    End If

    If pUseBrackets And cUseLinkIcons And Not vLinkedImage Then
        Dim vScheme, vImg, vPos
        vPos = Instr(pURL, ":")
        vScheme = Left(pURL, vPos - 1)
        vImg = "/wiki-" & vScheme & ".gif"" width=""12"" height=""12"""
        vLink = vLink & "<img src=""" & OPENWIKI_ICONPATH & vImg & " border=""0"" hspace=""4"" alt=""""/>" & pText
    Else
        If vLinkedImage Then
            vLink = vLink & pText
        Else
            If pUseBrackets And cShowBrackets Then
                vLink = vLink & "["
            End If
            vLink = vLink & pText
            If pUseBrackets And cShowBrackets Then
                vLink = vLink & "]"
            End If
        End If
    End If
    vLink = vLink & "</a>"
    GetExternalLink = vLink
End Function

Function GetBracketUrlIndex(pID)
    Dim i, vCount
    vCount = gBracketIndices.Count
    For i = 0 To vCount
        If gBracketIndices.ElementAt(i) = pID Then
            GetBracketUrlIndex = i + 1
            Exit Function
        End If
    Next
    gBracketIndices.Push(pID)
    GetBracketUrlIndex = gBracketIndices.Count
End Function

Sub UrlLink(pURL)
    Dim vLink, vTemp
    SplitUrlPunct(pURL)
    If cNetworkFile And (Left(pURL, 5) = "file:") Then
        ' only do remote file:// links. No file:///c|/windows.
        If (Left(pURL, 8) <> "file:///") Then
            gTempLink = "<a href=""" & gTempLink & """>" & gTempLink & "</a>"
        End If
        Exit Sub
    Elseif cAllowAttachments And (Left(pURL, 13) = "attachment://") Then
        gTempLink = AttachmentLink(Mid(gTempLink, 14), "")
        If gTempLink = "" Then
            gTempLink = pURL
        End If
        Exit Sub
    End If
    ' restricted image URLs so that mailto:foo@bar.gif is not an image
    If cLinkImages Then
        If m(gTempLink, gImagePattern, False, True) Then
            vLink = "<img src=""" & gTempLink & """ alt=""""/>"
        End If
    End If
    If vLink = "" Then
        vLink = "<a href=""" & gTempLink & """ class=""external"""
        If cExternalOut Then
            vLink = vLink & " target=""_blank"""
        End If
        vLink = vLink & ">" & gTempLink & "</a>"
    End If
    gTempLink = vLink
End Sub

Dim gTempLink, gTempJunk
Sub SplitUrlPunct(pURL)
    If Len(pURL) > 2 Then
        If Right(pURL, 2) = """""" Then
            gTempLink = Mid(pURL, 1, Len(pURL) - 2)
            gTempJunk = ""
            Exit Sub
        End If
    End If

    gTempLink = s(pURL, "([^a-zA-Z0-9\/\xc0-\xff]+)$", "", False, True)
    gTempJunk = Mid(pURL, Len(gTempLink) + 1)

    'Response.Write("GOT: " & Server.HTMLEncode(gTempLink) & "  :  " & Server.HTMLEncode(gTempJunk)& "<br>")

    ' check the rare case where a semicolon was actually part of the link
    ' e.g. http://x.com?x=<y> is, at this point, translated to <a ...>http://x.com?x=&lt;y&gt</a>;
    ' which is invalid XML
    If Left(gTempJunk, 1) = ";" Then
        gTemp = InStrRev(gTempLink, "&")
        If gTemp > 0 Then
            Dim vPosSemiColon
            vPosSemiColon = InStrRev(gTempLink, ";")
            If vPosSemiColon < gTemp Then
                ' invalid XML, restore
                gTempLink = gTempLink & ";"
                gTempJunk = Mid(gTempJunk, 2)
            End If
        End If
    End If
End Sub

Function AttachmentLink(pName, pText)
    Dim vPos, vPagename, vPage, vAttachment, vText
    If pText = "" Then
        vText = pName
    Else
        vText = Trim(pText)
    End If
    vPos = InStrRev(pName, "/")
    If vPos > 1 Then
        vPagename = Left(pName, vPos - 1)
        pName = Mid(pName, vPos + 1)
    Elseif IsObject(gCurrentWorkingPages) Then
        ' we're including a page
        vPagename = gCurrentWorkingPages.Top()
    Else
        vPagename = gPage
    End If

    Set vPage = gNamespace.GetPageAndAttachments(vPagename, gRevision, True, False)
    Set vAttachment = vPage.GetAttachment(pName)
    If vAttachment Is Nothing Then
        AttachmentLink = ""
        'AttachmentLink = "<ow:link name='" & CDATAEncode(pName) & "'" _
        '     & " href='" & gScriptName & "?p=" & Server.URLEncode(gPage) & "&amp;a=attach'" _
        '     & " attachment='true'>" _
        '     & PCDATAEncode(vText) & "</ow:link>"
    Elseif vAttachment.Deprecated Then
        AttachmentLink = ""
    Else
        AttachmentLink = vAttachment.ToXML(vPagename, vText)
    End If
End Function


Function InsertFootnotes(pText)
    pText = s(pText, gFS & gFS & "(.*?)" & gFS & gFS, "&AddFootnote($1)", False, True)
    If IsObject(gFootnotes) Then
        Dim i, vCount
        pText = pText & "<ow:footnotes>"
        For i = 0 To gFootnotes.Count - 1
            pText = pText & "<ow:footnote index='" & (i + 1) & "'>" & gFootnotes.ElementAt(i) & "</ow:footnote>"
        Next
        pText = pText & "</ow:footnotes>"
        Set gFootnotes = Nothing
    End If
    InsertFootnotes = pText
End Function

Dim gFootnotes
Sub AddFootnote(pParam)
    If Not IsObject(gFootnotes) Then
        Set gFootnotes = New Vector
    End If
    gFootnotes.Push(pParam)
    sReturn = "<sup><a href='#footnote" & gFootnotes.Count & "' class='footnote'>" & gFootnotes.Count & "</a></sup>"
End sub
%>