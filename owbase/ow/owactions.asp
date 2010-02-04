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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owactions.asp,v $
'    $Revision: 1.4 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

Sub ActionXml
    ActionView()
End Sub

Sub ActionRss
    Dim vPage, vXmlStr
    If cAllowRSSExport Then
        If Request("p") <> "" And cAllowAggregations Then
            Set vPage = gNamespace.GetPage(gPage, gRevision, True, False)
            Set gAggregateURLs = New Vector
            Set gRaw = New Vector
            MultiLineMarkup(vPage.Text)   ' refreshes RSS feed(s) and fills the gAggregateURLs vector
            If gAggregateURLs.Count = 0 Then
                Response.ContentType = "text/xml; charset:" & OPENWIKI_ENCODING & ";"
                Response.Write "<?xml version='1.0'?><error>Nothing to aggregate</error>"
                Response.End
            Else
                Response.ContentType = "text/xml; charset:" & OPENWIKI_ENCODING & ";"
                Response.Write gNamespace.GetAggregation(gAggregateURLs)
                Response.End
            End If
        Else
            If cCacheXML Then
                vXmlStr = gNamespace.GetDocumentCache("rss")
            End If
            If vXmlStr = "" Then
                gPage = OPENWIKI_RCNAME
                Set vPage = gNamespace.GetPage(gPage, gRevision, False, False)
                ' make sure we execute only the RecentChanges macro
                vPage.Text = "<RecentChangesLong>"
                vXmlStr = gTransformer.TransformXmlStr(vPage.ToXML(1), "owrss10export.xsl")
                If cCacheXML Then
                    Call gNamespace.SetDocumentCache("rss", vXmlStr)
                End If
            End If
            gActionReturn = True
        End If
    Else
        Response.ContentType = "text/xml; charset:" & OPENWIKI_ENCODING & ";"
        Response.Write "<?xml version='1.0'?><error>RSS feed disabled</error>"
        Response.End
    End If
End Sub

Sub ActionRefresh
    Dim vPage
    If OPENWIKI_SCRIPTTIMEOUT > 0 Then
        Server.ScriptTimeout = OPENWIKI_SCRIPTTIMEOUT
    End If
    cCacheXML = False
    Set vPage = gNamespace.GetPage(gPage, gRevision, True, False)
    Set gAggregateURLs = New Vector
    Set gRaw = New Vector
    Call MultiLineMarkup(vPage.Text)   ' refreshes RSS feed(s)
    Call gNamespace.ClearDocumentCache2("", gPage)
    If Request("redirect") = "" Then
        Response.Redirect(gScriptName & "?" & Server.URLEncode(gPage))
    Else
        Response.Redirect(gScriptName & "?" & Server.URLEncode(Request("redirect")))
    End If
End Sub

Sub ActionNaked
    gAction = "view"
    ActionView()
End Sub

Sub ActionPrint
    Dim vXmlStr
    cReadOnly = 1
    If cCacheXML Then
        vXmlStr = gNamespace.GetDocumentCache("print")
    End If
    If vXmlStr = "" Then
        vXmlStr = gNamespace.GetPageAndAttachments(gPage, gRevision, True, False).ToXML(1)
        If cCacheXML Then
            Call gNamespace.SetDocumentCache("print", vXmlStr)
        End If
    End If
    Call gTransformer.Transform(vXmlStr)
    gActionReturn = True
End Sub

Sub ActionView
    Dim vXmlStr
    If cNakedView Then
        gAction = "naked"
    End If
    If cAllowRSSExport And Request("v") = "rss" Then
        Call gTransformer.TransformXmlStr(gNamespace.GetPage(gPage, gRevision, True, False).ToXML(1), "owrss10export.xsl")
    Else
        If cCacheXML Then
            vXmlStr = gNamespace.GetDocumentCache("view")
        End If
        If vXmlStr = "" Then
            vXmlStr = gNamespace.GetPageAndAttachments(gPage, gRevision, True, False).ToXML(1)
            If cCacheXML Then
                Call gNamespace.SetDocumentCache("view", vXmlStr)
            End If
        End If
        Call gTransformer.Transform(vXmlStr)
    End If
    gActionReturn = True
End Sub

Sub ActionPreview
    Dim vPage
    Set vPage = gNamespace.GetPage(gPage, 0, False, False)
    vPage.Text = Request("text")
    gAction = "print"
    Call gTransformer.Transform(vPage.ToXML(1))
    gActionReturn = True
End Sub

Sub ActionDiff
    Dim vXmlStr, vDiff, vDiffFrom, vDiffTo, vDiffType, vPageFrom, vPageTo, vMatcher
    vDiff     = GetIntParameter("diff")
    vDiffFrom = GetIntParameter("difffrom")
    vDiffTo   = GetIntParameter("diffto")

    If vDiffFrom <> 0 Or vDiffTo <> 0 Then
        cCacheXML = False
    End If

    If cCacheXML Then
        vXmlStr = gNamespace.GetDocumentCache("diff" & vDiff)
    End If

    If vXmlStr = "" Then

        If vDiff = 0 Then
            If vDiffFrom = 0 Then
                ' difference of prior major revision relative to vDiffTo
                vDiffType = "major"
                vDiffFrom = gNamespace.GetPreviousRevision(0, vDiffTo)
            Else
                ' difference of selected revision relative to vDiffTo
                vDiffType = "selected"
            End If
        Elseif vDiff = 1 Then
            ' difference of previous minor edit relative to vDiffTo
            vDiffType = "minor"
            vDiffFrom = gNamespace.GetPreviousRevision(1, vDiffTo)
        Else
            ' difference of previous author edit relative to vDiffTo
            vDiffType = "author"
            vDiffFrom = gNamespace.GetPreviousRevision(2, vDiffTo)
        End If

        ' difference of vDiffFrom to vDiffTo
        Set vPageFrom = gNamespace.GetPage(gPage, vDiffFrom, True, False)
        Set vPageTo   = gNamespace.GetPageAndAttachments(gPage, vDiffTo, True, False)
        vDiffFrom = vPageFrom.GetLastChange().Revision
        vDiffTo   = vPageTo.GetLastChange().Revision
        vXmlStr = "<ow:diff type='" & vDiffType & "' from='" & vDiffFrom & "' to='" & vDiffTo & "'>"
        If vDiffTo > vDiffFrom Then
            Set vMatcher = New Matcher
            vXmlStr = vXmlStr & vMatcher.Compare(Server.HTMLEncode(vPageFrom.Text), Server.HTMLEncode(vPageTo.Text))
        End If
        vXmlStr = vXmlStr & "</ow:diff>"
        vXmlStr = vXmlStr & vPageTo.ToXML(1)

        If cCacheXML Then
            Call gNamespace.SetDocumentCache("diff" & vDiff, vXmlStr)
        End If
    End If

    Call gTransformer.Transform(vXmlStr)
    Set vMatcher  = Nothing
    Set vPageTo   = Nothing
    Set vPageFrom = Nothing

    gActionReturn = True
End Sub

Sub ActionEdit
    Dim vPage, vChange, vXmlStr
    Dim vNewRev, vMinorEdit, vComment, vText

    If cReadOnly Then
        ' TODO: generate <ow:error> tag into the XML output
        gAction = "view"
        ActionView()
        gActionReturn = True
        Exit Sub
    End If

    If gEditPassword <> "" Then
        If gEditPassword <> gReadPassword Then
            If Request.Cookies(gCookieHash & "?pe") <> gEditPassword Then
                Call ActionLogin()
                Exit Sub
            End If
        End If
    End If

    If Request("save") <> "" Then
        vNewRev    = Int(Request("newrev"))
        vMinorEdit = Int(Request("rc")) Xor 1
        vComment   = Trim(Request("comment") & "")
        vText      = Request("text")

        If Len(vComment) > 1000 Then
            vXmlStr = vXmlStr & "<ow:error code='1'>Maximum length for the comment is 1000 characters.</ow:error>"
        End If
        If Len(vText) > OPENWIKI_MAXTEXT Then
            vXmlStr = vXmlStr & "<ow:error code='2'>Maximum length for the text is " & OPENWIKI_MAXTEXT & " characters.</ow:error>"
        End If

        If vXmlStr <> "" Then
            Set vPage = gNamespace.GetPage(gPage, 0, False, False)
            vPage.Revision = gRevision
            vPage.Text     = vText

            Set vChange = vPage.GetLastChange()
            vChange.Revision  = vNewRev
            vChange.MinorEdit = vMinorEdit
            vChange.Comment   = vComment
            vChange.Timestamp = Now()
            vChange.UpdateBy()

            vXmlStr = vXmlStr & vPage.ToXML(2)
        Elseif gNamespace.SavePage(vNewRev, vMinorEdit, vComment, vText) Then
            Response.Redirect(gScriptName & "?" & Server.URLEncode(gPage))
        Else
            Set vPage = gNamespace.GetPage(gPage, 0, True, False)
            Set vChange = vPage.GetLastChange()
            vChange.Revision  = vChange.Revision + 1
            vChange.MinorEdit = Int(Request("rc")) Xor 1
            vChange.Comment   = Trim(Request("comment") & "")
            vChange.Timestamp = Now()
            vChange.UpdateBy()
            vXmlStr = vXmlStr & "<ow:error code='4'>Somebody else just edited this page.</ow:error>"
            vXmlStr = vXmlStr & "<ow:textedits>" & PCDATAEncode(Request("text")) & "</ow:textedits>"
            vXmlStr = vXmlStr & vPage.ToXML(2)
        End If
    ' now v0.78, let's see if someone's going to complain..
    'Elseif Request("preview") <> "" Then
        ' pre 0.74 version code; now ActionPreview (i.e. ?a=preview is prefered method)

    '    vNewRev    = Int(Request("newrev"))
    '    vMinorEdit = Int(Request("rc")) Xor 1
    '    vComment   = Trim(Request("comment") & "")
    '    vText      = Request("text")
    '
    '    Set vPage = gNamespace.GetPage(gPage, 0, False, False)
    '    vPage.Revision = gRevision
    '    vPage.Text     = vText
    '
    '    Set vChange = vPage.GetLastChange()
    '    vChange.Revision  = vNewRev
    '    vChange.MinorEdit = vMinorEdit
    '    vChange.Comment   = vComment
    '    vChange.Timestamp = Now()
    '    vChange.UpdateBy()
    '
    '    vXmlStr = vPage.ToXML(3)
    Elseif Request("cancel") <> "" Then
        Dim vBacklink
        If gRevision = 0 Then
            vBacklink = gScriptName & "?" & Server.URLEncode(gPage)
        Else
            vBacklink = gScriptName & "?p=" & Server.URLEncode(gPage) & "&revision=" & gRevision
        End If
        Response.Redirect(vBacklink)
    Else
        ' first time opening edit form
        Set vPage = gNamespace.GetPage(gPage, 0, True, False)
        If gRevision > 0 Then
            Set gTemp = gNamespace.GetPage(gPage, gRevision, True, False)
            vPage.Revision = gTemp.Revision
            vPage.Text     = gTemp.Text
        End If

        If vPage.Revision = 0 And Request("template") <> "" Then
            Set gTemp = gNamespace.GetPage(URLDecode(Request("template")), 0, True, False)
            vPage.Text = gTemp.Text
        End If

        Set vChange = vPage.getLastChange()
        vChange.Revision  = vChange.Revision + 1
        vChange.MinorEdit = 0
        vChange.Comment   = ""
        vChange.Timestamp = Now()
        vChange.UpdateBy()

        vXmlStr = vPage.ToXML(2)
    End If

    Call gTransformer.Transform(vXmlStr)
    gActionReturn = True
End Sub

Sub ActionTitleSearch
    Dim vXmlStr
    vXmlStr = gNamespace.GetIndexSchemes.GetTitleSearch(gTxt)
    If cAllowRSSExport And Request("v") = "rss" Then
        Call gTransformer.TransformXmlStr(vXmlStr, "owsearchrss10export.xsl")
    Else
        Call gTransformer.Transform(vXmlStr)
    End If
    gActionReturn = True
End Sub

Sub ActionFullSearch
    Dim vXmlStr
    vXmlStr = gNamespace.GetIndexSchemes.GetFullSearch(gTxt, True)
    If cAllowRSSExport And Request("v") = "rss" Then
        Call gTransformer.TransformXmlStr(vXmlStr, "owsearchrss10export.xsl")
    Else
        Call gTransformer.Transform(vXmlStr)
    End If
    gActionReturn = True
End Sub

Sub ActionTextSearch
    Dim vXmlStr
    vXmlStr = gNamespace.GetIndexSchemes.GetFullSearch(gTxt, False)
    If cAllowRSSExport And Request("v") = "rss" Then
        Call gTransformer.TransformXmlStr(vXmlStr, "owsearchrss10export.xsl")
    Else
        Call gTransformer.Transform(vXmlStr)
    End If
    gActionReturn = True
End Sub

Sub ActionRandomPage
    Randomize
    Set gTemp = gNamespace.TitleSearch(".*", 0, 0, 0, 0)
    Response.Redirect(gScriptName & "?a=" & gAction & "&p=" & Server.URLEncode(gTemp.ElementAt(Int((gTemp.Count - 1) * Rnd)).Name) & "&redirect=" & Server.URLEncode(gPage))
End Sub

Sub ActionChanges
    Dim vXmlStr
        If cCacheXML Then
            vXmlStr = gNamespace.GetDocumentCache("changes")
        End If
        If vXmlStr = "" Then
            vXmlStr = gNamespace.GetPage(gPage, 0, False, True).ToXML(0)
            If cCacheXML Then
                Call gNamespace.SetDocumentCache("changes", vXmlStr)
            End If
        End If
        Call gTransformer.Transform(vXmlStr)
    gActionReturn = True
End Sub

Sub ActionUserPreferences
    If Request("save") <> "" Then
        Response.Cookies(gCookieHash & "?up").Expires  = Date + 60
        Response.Cookies(gCookieHash & "?up")("un")   = FreeToNormal(Request("username"))
        Response.Cookies(gCookieHash & "?up")("bm")   = Request("bookmarks")
        Response.Cookies(gCookieHash & "?up")("cols") = Request("cols")
        Response.Cookies(gCookieHash & "?up")("rows") = Request("rows")
        Response.Cookies(gCookieHash & "?up")("pwl")  = Request("prettywikilinks")
        Response.Cookies(gCookieHash & "?up")("bmt")  = Request("bookmarksontop")
        Response.Cookies(gCookieHash & "?up")("elt")  = Request("editlinkontop")
        Response.Cookies(gCookieHash & "?up")("trt")  = Request("trailontop")
        Response.Cookies(gCookieHash & "?up")("new")  = Request("opennew")
        Response.Cookies(gCookieHash & "?up")("emo")  = Request("emoticons")
        Response.Redirect(gScriptName & "?p=" & Server.URLEncode(gPage) & "&up=1")
    Elseif Request("clear") <> "" Then
        Response.Cookies(gCookieHash & "?up").expires  = #01/01/1990#
        Response.Cookies(gCookieHash & "?up") = ""
        Response.Redirect(gScriptName & "?p=" & Server.URLEncode(gPage) & "&up=2")
    End If
    gActionReturn = False
End Sub

Sub ActionLogout
     Response.Cookies(gCookieHash & "?pr").Expires = #01/01/1990#
     Response.Cookies(gCookieHash & "?pr") = ""
     Response.Cookies(gCookieHash & "?pe").Expires = #01/01/1990#
     Response.Cookies(gCookieHash & "?pe") = ""
     Response.Redirect(gScriptName & "?" & Server.URLEncode(gPage))
End Sub

Sub ActionLogin
    Dim vMode, vPwd, vXmlStr
    If gAction = "edit" Then
        vMode = "edit"
        gAction = "login"
    Else
        vMode = Request("mode")
    End If
    vPwd = Request("pwd")
    If vMode = "edit" Then
        If vPwd = gEditPassword Then
            If Request("r") = "1" Then
                Response.Cookies(gCookieHash & "?pe").Expires = Date + 60
            End If
            Response.Cookies(gCookieHash & "?pe") = vPwd
            Response.Redirect(gScriptName & "?" & Request("backlink"))
        End If
    Else
        If vPwd = gReadPassword Then
            If Request("r") = "1" Then
                Response.Cookies(gCookieHash & "?pr").Expires = Date + 60
            End If
            Response.Cookies(gCookieHash & "?pr") = vPwd
            Response.Redirect(gScriptName & "?" & Request("backlink"))
        End If
    End If
    If vPwd <> "" Then
        vXmlStr = "<ow:error code='3'>Incorrect password</ow:error>"
    End If
    If Request("backlink") <> "" Then
        gTemp = Request("backlink")
    Else
        gTemp = Request.ServerVariables("QUERY_STRING")
        If gTemp = "" Then
            gTemp = OPENWIKI_FRONTPAGE
        End If
    End If
    vXmlStr = vXmlStr & "<ow:login"
    If vMode = "edit" Then
        vXmlStr = vXmlStr & " mode='edit'>"
    Else
        vXmlStr = vXmlStr & " mode='view'>"
    End If
    vXmlStr = vXmlStr & "<ow:backlink>" & PCDATAEncode(gTemp) & "</ow:backlink>"
    If Request("r") <> "" Then
        vXmlStr = vXmlStr & "<ow:rememberme>true</ow:rememberme>"
    End If
    vXmlStr = vXmlStr & "</ow:login>"
    Call gTransformer.Transform(vXmlStr)
    gActionReturn = True
End Sub


%>

