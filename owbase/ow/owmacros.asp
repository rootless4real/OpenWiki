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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owmacros.asp,v $
'    $Revision: 1.4 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

Sub ExecMacro(pMacro, pParams)
    ' On error resume next should be on, because in the event someone does e.g. <bogusmacroname>
    ' then it should nicely return.
    ' The side effect of having this option on is that if a programming error occurs in the
    ' processing of a macro, the programmer won't notice it.
    On Error Resume Next
    Dim vMacro, vParams, vPos, vTemp1, vTemp2, vCmd
    vMacro  = pMacro
    vParams = pParams
    If vParams <> "" Then
        If IsNumeric(vParams) Then
            If InStr(vParams, ",") > 0 Then
                vMacro = vMacro & "P"
            End If
        Else
            If Mid(vParams, 2, 1) = """" Then
                vPos = InStr(3, vParams, """")
                If InStr(vPos, vParams, ",") > 0 Then
                    vMacro = vMacro & "P"
                End If
            Else
                vPos = InStr(vParams, ",")
                If vPos > 0 Then
                    vTemp1 = Mid(vParams, 2, vPos - 2)
                    If Not IsNumeric(vTemp1) Then
                        vTemp1 = """" & vTemp1 & """"
                    End If
                    vTemp2 = Mid(vParams, vPos + 1, Len(vParams) - vPos - 1)
                    If Not IsNumeric(vTemp2) Then
                        vTemp2 = """" & vTemp2 & """"
                    End If
                    vParams = "(" & vTemp1 & "," & vTemp2 & ")"
                    vMacro = vMacro & "P"
                Else
                    vParams = "(""" & Mid(vParams, 2, Len(vParams) - 2) & """)"
                End If
            End If
        End If
        vMacro = vMacro & "P"
    End If

    gMacroReturn = ""
    vCmd = "Macro" & vMacro & vParams
    vCmd = Replace(vCmd, vbCRLF, """ & vbCRLF & """)
    'Response.Write("<br />MACRO CMD: " & Server.HTMLEncode(vCmd))
    Execute("Call " & vCmd)
    If gMacroReturn = "" Then
        sReturn = "&lt;" & pMacro & pParams & "&gt;"
    Else
        StoreRaw(gMacroReturn)
    End If
End Sub

Sub MacroTableOfContents
    If cUseHeadings Then
        gMacroReturn = gFS & "TOC" & gFS
        ' at the end of the Wikify function this pattern will be
        ' replaced by the actual table of contents
    End If
End Sub

Sub MacroBR
    gMacroReturn = "<br />"
End Sub

Sub MacroTitleSearch
    gMacroReturn = "<form name=""TitleSearch"" action=""" & CDATAEncode(gScriptName) & """ method=""get""><input type=""hidden"" name=""a"" value=""titlesearch""/><input type=""text"" name=""txt"" value=""" & CDATAEncode(gTxt) & """ ondblclick='event.cancelBubble=true;' /><input id=""mts"" type=""submit"" value=""Go""/></form>"
End Sub

Sub MacroTitleSearchP(pParam)
    gMacroReturn = gNamespace.GetIndexSchemes.GetTitleSearch(pParam)
End Sub

Sub MacroFullSearch
    gMacroReturn = "<form name=""FullSearch"" action=""" & CDATAEncode(gScriptName) & """ method=""get""><input type=""hidden"" name=""a"" value=""fullsearch""/><input type=""text"" name=""txt"" value=""" & CDATAEncode(gTxt) & """ ondblclick='event.cancelBubble=true;' /><input id=""mfs"" type=""submit"" value=""Go""/></form>"
End Sub

Sub MacroFullSearchP(pParam)
    gMacroReturn = gNamespace.GetIndexSchemes.GetFullSearch(pParam, True)
End Sub

Sub MacroTextSearch
    gMacroReturn = "<form name=""TextSearch"" action=""" & CDATAEncode(gScriptName) & """ method=""get""><input type=""hidden"" name=""a"" value=""textsearch""/><input type=""text"" name=""txt"" value=""" & CDATAEncode(gTxt) & """ ondblclick='event.cancelBubble=true;' /><input id=""mfs"" type=""submit"" value=""Go""/></form>"
End Sub

Sub MacroTextSearchP(pParam)
    gMacroReturn = gNamespace.GetIndexSchemes.GetFullSearch(pParam, False)
End Sub

Sub MacroGoTo
    gMacroReturn = "<form name=""GoTo"" action=""" & CDATAEncode(gScriptName) & """ method=""get""><input type=""text"" name=""p"" value="""" ondblclick='event.cancelBubble=true;' /><input id=""goto"" type=""submit"" value=""Go""/></form>"
End Sub

Sub MacroSystemInfo
    On Error Resume Next
    Dim vFSO, vFile
    Set vFSO = server.CreateObject("Scripting.FileSystemObject")
    Set vFile = vFSO.GetFile(Server.MapPath(Request.ServerVariables("SCRIPT_NAME")))
    Dim vRev
    vRev = Mid(OPENWIKI_REVISION, 12, Len(OPENWIKI_REVISION) - 13)
    gMacroReturn = "<table class=""systeminfo"">" _
            & "<tr><td>OpenWiki Version:</td><td>" & OPENWIKI_VERSION & " rev." & vRev & "</td></tr>" _
            & "<tr><td>XML Schema Version:</td><td>" & OPENWIKI_XMLVERSION & "</td></tr>" _
            & "<tr><td>Namespace:</td><td>" & OPENWIKI_NAMESPACE & "</td></tr>" _
            & "<tr><td>" & ScriptEngine & " Version:</td><td>" & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion & "</td></tr>"
    Dim vConn
    Set vConn = Server.CreateObject("ADODB.Connection")
    gMacroReturn = gMacroReturn & "<tr><td>ADO Version:</td><td>" & vConn.Version & "</td></tr>"
    Set vFile = Nothing
    Set vFSO = Nothing
    gMacroReturn = gMacroReturn & "<tr><td>Nr Of Pages:</td><td>" & gNamespace.GetPageCount() & "</td></tr>"
    gMacroReturn = gMacroReturn & "<tr><td>Nr Of Revisions:</td><td>" & gNamespace.GetRevisionsCount() & "</td></tr>"
    'gMacroReturn = gMacroReturn & "<tr><td>Now:</td><td>" & FormatDate(Now()) & " " & FormatTime(Now()) & "</td></tr>"
    gMacroReturn = gMacroReturn & "</table>"
End Sub

Sub MacroDate
    cCacheXML = False
    gMacroReturn = FormatDate(Now())
End Sub

Sub MacroTime
    cCacheXML = False
    gMacroReturn = FormatTime(Now())
End Sub

Sub MacroDateTime
    cCacheXML = False
    gMacroReturn = FormatDate(Now()) & " " & FormatTime(Now())
End Sub

Sub MacroPageCount
    gMacroReturn = gNamespace.GetPageCount()
End Sub

Sub MacroRecentChanges
    Call MacroRecentChangesPP(OPENWIKI_RCDAYS, 9999)
End Sub

Sub MacroRecentChangesP(pParams)
    Call MacroRecentChangesPP(pParams, 9999)
End Sub

Sub MacroRecentChangesPP(pDays, pNrOfChanges)
    If Not IsNumeric(pDays) Or Not IsNumeric(pNrOfChanges) Then
        Exit Sub
    End If
    If pDays <= 0 Then
        pDays = OPENWIKI_RCDAYS
    End If
    If pNrOfChanges <= 0 Then
        pNrOfChanges = 0
    End If
    gMacroReturn = gNamespace.GetIndexSchemes.GetRecentChanges(pDays, pNrOfChanges, 1, True)
End Sub

Sub MacroRecentChangesLong
    Dim vDays, vMaxNrOfChanges, vFilter
    vDays = GetIntParameter("days")
    vMaxNrOfChanges = GetIntParameter("max")
    vFilter = GetIntParameter("filter")
    If vDays <= 0 Then
        vDays = OPENWIKI_RCDAYS
    End If
    If vMaxNrOfChanges <= 0 Then
        If gAction = "rss" Then
            vMaxNrOfChanges = 15
        Else
            vMaxNrOfChanges = 9999
        End If
    End If
    If vFilter = 0 Then
        vFilter = 1  ' major edits only
    Elseif vFilter = 3 Then
        vFilter = 0  ' major and minor edits
    End If
    ' vFilter = 2  ' minor edits only
    gMacroReturn = gNamespace.GetIndexSchemes.GetRecentChanges(vDays, vMaxNrOfChanges, vFilter, False)
End Sub

Sub MacroTitleIndex
    gMacroReturn = gNamespace.GetIndexSchemes.GetTitleIndex
End Sub

Sub MacroWordIndex
    gMacroReturn = gNamespace.GetIndexSchemes.GetWordIndex
End Sub

Sub MacroRandomPage
    gMacroReturn = gNamespace.GetIndexSchemes.GetRandomPage(1)
End Sub

Sub MacroRandomPageP(pParam)
    If IsNumeric(pParam) Then
        gMacroReturn = gNamespace.GetIndexSchemes.GetRandomPage(pParam)
    End If
End Sub

Sub MacroIconP(pParam)
    gMacroReturn = "<img src=""" & OPENWIKI_ICONPATH & "/" & pParam & ".gif"" border=""0"" alt=""" & pParam & """/>"
End Sub

Sub MacroAnchorP(pParam)
    gMacroReturn = "<a name='" & CDATAEncode(pParam) & "'></a>"
End Sub

Sub MacroIncludeP(pParam)
    Dim i, vCount, vID
    If Not IsObject(gCurrentWorkingPages) Then
        Set gCurrentWorkingPages = New Vector
        gCurrentWorkingPages.Push(gPage)
    End If
    For i = 0 To gCurrentWorkingPages.Count - 1
        If UCase(gCurrentWorkingPages.ElementAt(i)) = UCase(pParam) Then
            Exit Sub
        End If
    Next

    vID = AbsoluteName(pParam)

    gIncludeLevel = gIncludeLevel + 1
    If (gIncludeLevel <= OPENWIKI_MAXINCLUDELEVEL) Then
        Dim vPage
        Set vPage = gNamespace.GetPageAndAttachments(vID, 0, True, False)
        If vPage.Exists Then
            gCurrentWorkingPages.Push(vPage.Name)
            gMacroReturn = vPage.ToXML(1)
            gCurrentWorkingPages.Pop()
        End If
    End If
    gIncludeLevel = gIncludeLevel - 1
End Sub

Sub MacroInterWiki
    gMacroReturn = gNamespace.InterWiki()
End Sub

Sub MacroUserPreferences
    gMacroReturn = ""
    If Request.QueryString("up") = 1 Then
        gMacroReturn = gMacroReturn & "<ow:message code=""userpreferences_saved""/>"
    Elseif Request.QueryString("up") = 2 Then
        gMacroReturn = gMacroReturn & "<ow:message code=""userpreferences_cleared""/>"
    End If
    gMacroReturn = gMacroReturn & "<ow:userpreferences/>"
End Sub

Function FormatDate(pTimestamp)
    ' TODO: apply user preferences
    FormatDate = MonthName(Month(pTimestamp)) & " " & Day(pTimestamp) & ", " & Year(pTimestamp)
End Function

Function FormatTime(pTimestamp)
    ' TODO: apply user preferences
    FormatTime = FormatDateTime(pTimestamp, 4)  ' 4 = vbShortTime
End Function


Sub MacroFootnoteP(pText)
    ' processed at the end of wikify
    gMacroReturn = gFS & gFS & pText & gFS & gFS
End Sub


Sub MacroAggregateP(pPage)
    If cAllowAggregations <> 1 Then
        Exit Sub
    End If

    If Request("preview") <> "" Then
        Exit Sub
    End If

    pPage = AbsoluteName(pPage)

    Dim vPage
    Set vPage = gNamespace.GetPage(pPage, gRevision, True, False)
    Set gAggregateURLs = New Vector
    MultiLineMarkup(vPage.Text)   ' refreshes RSS feed(s) and fills the gAggregateURLs vector
    gMacroReturn = GetAggregation(pPage)
    gAggregateURLs = ""
End Sub

Sub MacroSyndicateP(pURL)
    Call MacroSyndicatePP(pURL, 240)  ' default = 4 * 60 minutes
End Sub

Sub MacroSyndicatePP(pURL, pRefreshRate)
    Dim vURL, vCache, vRefreshURL

    If Request("preview") <> "" Then
        Exit Sub
    End If

    vURL = Replace(pURL, "&amp;", "&")
    If Not m(vURL, "^https?://", False, False) Or Not IsNumeric(pRefreshRate) Then
        Exit Sub
    End If
    If pRefreshRate < 0 Then
        pRefreshRate = 0
    End If

    If IsObject(gAggregateURLs) And cAllowAggregations Then
        gAggregateURLs.Push(vURL)
    End If

    vRefreshURL = URLDecode(Request("refreshurl"))

    If (gAction <> "refresh") Or ((vRefreshURL <> "") And (vRefreshURL <> vURL)) Then
        vCache = gNamespace.GetRSSFromCache(vURL, pRefreshRate, False, False)
        If vCache = "notexists" Then
            If cAllowNewSyndications = 0 Then
                Exit Sub
            End If
        Elseif vCache <> "" Then
            gMacroReturn = vCache
            Exit Sub
        End If
    End If
    If gAction = "refresh" Or vRefreshURL = vURL Or vCache = "notexists" Or gNrOfRSSRetrievals < OPENWIKI_MAXWEBGETS Then
        gMacroReturn = RetrieveRSSFeed(vURL)
        gNrOfRSSRetrievals = gNrOfRSSRetrievals + 1
    End If
    If gMacroReturn = "" Then
        ' failure to retrieve RSS feed from remote source
        If vCache = "notexists" Then
            Call gNamespace.SaveRSSToCache(vURL, pRefreshRate, "")
            gMacroReturn = gNamespace.GetRSSFromCache(vURL, pRefreshRate, True, False)
        Else
            ' retry later, and get the cached version
            gMacroReturn = gNamespace.GetRSSFromCache(vURL, pRefreshRate, True, True)
            If gMacroReturn = "notexists" Then
                gMacroReturn = ""
            End If
        End If
    Else
        Call gNamespace.SaveRSSToCache(vURL, pRefreshRate, gMacroReturn)
        gMacroReturn = gNamespace.GetRSSFromCache(vURL, pRefreshRate, True, False)
    End If
End Sub
%>
