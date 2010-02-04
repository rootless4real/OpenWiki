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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owdb.asp,v $
'    $Revision: 1.6 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

Class OpenWikiNamespace
    Private vConn, vRS, vQuery
    Private vIndexSchemes
    Private vCachedPages

    Private Sub Class_Initialize()
        If OPENWIKI_DB = "" Then
            cAllowAttachments = 0
            cWikiLinks = 0
            cCacheXML = 0
        Else
            Set vConn = Server.CreateObject("ADODB.Connection")
            vConn.Open OPENWIKI_DB
            Set vRS = Server.CreateObject("ADODB.Recordset")
        End If
        Set vIndexSchemes = New IndexSchemes
        Set vCachedPages = Server.CreateObject("Scripting.Dictionary")
    End Sub

    Private Sub Class_Terminate()
        On Error Resume Next
        vConn.Close
        Set vConn = Nothing
        Set vRS = Nothing
        Set vIndexSchemes = Nothing
        Set vCachedPages = Nothing
    End Sub

    Sub BeginTrans(pConn)
        If OPENWIKI_DB_SYNTAX <> DB_MYSQL Then
            pConn.BeginTrans()
        End If
    End Sub

    Sub CommitTrans(pConn)
        If OPENWIKI_DB_SYNTAX <> DB_MYSQL Then
            pConn.CommitTrans()
        End If
    End Sub

    Sub RollbackTrans(pConn)
        If OPENWIKI_DB_SYNTAX <> DB_MYSQL Then
            pConn.RollbackTrans()
        End If
    End Sub

    Private Function CreatePageKey(pPageName, pRevision, pIncludeText, pIncludeAllChangeRecords)
        CreatePageKey = pRevision & "_" & pIncludeText & "_" & pIncludeAllChangeRecords & "_" & pPageName
    End Function

    Private Function GetCachedPage(pPageName, pRevision, pIncludeText, pIncludeAllChangeRecords)
        Dim vKey
        vKey = CreatePageKey(pPageName, pRevision, pIncludeText, pIncludeAllChangeRecords)
        If vCachedPages.Exists(vKey) Then
            Set GetCachedPage = vCachedPages.Item(vKey)
        Else
            Set GetCachedPage = Nothing
        End If
    End Function

    Private Sub SetCachedPage(pPageName, pRevision, pIncludeText, pIncludeAllChangeRecords, vPage)
        Dim vKey
        vKey = CreatePageKey(pPageName, pRevision, pIncludeText, pIncludeAllChangeRecords)
        vCachedPages.Add vKey, vPage
    End Sub

    Public Function GetIndexSchemes
        Set GetIndexSchemes = vIndexSchemes
    End Function

    Function GetPageAndAttachments(pPageName, pRevision, pIncludeText, pIncludeAllChangeRecords)
        Dim vPage
        Set vPage = GetCachedPage(pPageName, pRevision, pIncludeText, pIncludeAllChangeRecords)
        If TypeName(vPage) = TypeName(Nothing) Then
            Set vPage = GetPage(pPageName, pRevision, pIncludeText, False)
            If cAllowAttachments Then
                Call GetAttachments(vPage, pRevision, pIncludeAllChangeRecords)
            End If
        Elseif cAllowAttachments And Not vPage.AttachmentsLoaded Then
            Call GetAttachments(vPage, pRevision, pIncludeAllChangeRecords)
        End If
        Set GetPageAndAttachments = vPage
    End Function

    Function GetPage(pPageName, pRevision, pIncludeText, pIncludeAllChangeRecords)
        Dim vPage, vChange
        If cWikiLinks = 0 Then
            Set GetPage = New WikiPage
            GetPage.AddChange
            GetPage.Name = "FrontPage"
            GetPage.Text = "Please provide a value for {{{OPENWIKI_DB}}} in your owconfig.asp file."
            Exit Function
        End If
        Set vPage = GetCachedPage(pPageName, pRevision, pIncludeText, pIncludeAllChangeRecords)
        If TypeName(vPage) = TypeName(Nothing) then
            'Response.Write("LOAD PAGE: " & pPageName & "<br />")
            Set vPage = new WikiPage
            If pIncludeText Then
                vQuery = "SELECT * "
            Else
                vQuery = "SELECT wpg_name, wpg_changes, wpg_lastminor, wpg_lastmajor, wrv_revision, wrv_status, wrv_timestamp, wrv_minoredit, wrv_by, wrv_byalias, wrv_comment "
            End If
            vQuery = vQuery & " FROM openwiki_pages, openwiki_revisions WHERE wpg_name = '" & Replace(pPageName, "'", "''") & "' AND wrv_name = wpg_name"
            If pRevision > 0 Then
                vQuery = vQuery & " AND wrv_revision = " & pRevision
            Elseif pIncludeAllChangeRecords Then
                vQuery = vQuery & " ORDER BY wrv_revision DESC"
            Else
                vQuery = vQuery & " AND wrv_current = 1"
            End If

            On Error Resume Next
            vRS.Open vQuery, vConn, adOpenForwardOnly
            If Err.Number <> 0 Then
                If Err.Number = -2147467259 Then
                    Response.Write("<h2>Error:</h2>")
                    Response.Write("Cannot find the data sources or the data sources are locked by another application.")
                    Response.Write("Make sure you've set the constant <code><b>OPENWIKI_DB</b></code> correctly in your config file, pointing it to your data sources.<br /><br /><br />")
                Else
                    Response.Write(Err.Number & "<br />" & Err.Description)
                End If
                Response.End
            End If
            On Error Goto 0

            If vRS.EOF Then
                If pRevision = 0 Then
                    vPage.Name = pPageName
                    vPage.AddChange
                Else
                    ' TODO: addMessage("Revision " & pRevision & " not available (showing current version instead)"
                    vRS.Close
                    Set GetPage = GetPage(pPageName, 0, pIncludeText, pIncludeAllChangeRecords)
                    Exit Function
                End If
            Else
                vPage.Name      = vRS("wpg_name")
                vPage.Changes   = CInt(vRS("wpg_changes"))
                vPage.LastMinor = CInt(vRS("wpg_lastminor"))
                vPage.LastMajor = CInt(vRS("wpg_lastmajor"))
                If pIncludeText Then
                    vPage.Text = vRS("wrv_text")
                End If
                If CInt(vRS("wpg_lastminor")) = CInt(vRS("wrv_revision")) Then
                    ' wrv_current = 1
                    ' vPage.Revision = vRS("wrv_revision") ??? ---> No! Because of the xsl script.
                    vPage.Revision = 0
                Elseif pRevision > 0 Then
                    vPage.Revision = pRevision
                End If
                Do While Not vRS.EOF
                    Set vChange = vPage.AddChange
                    vChange.Revision  = CInt(vRS("wrv_revision"))
                    vChange.Status    = CInt(vRS("wrv_status"))
                    vChange.MinorEdit = CInt(vRS("wrv_minoredit"))
                    vChange.Timestamp = vRS("wrv_timestamp")
                    vChange.By        = vRS("wrv_by")
                    vChange.ByAlias   = vRS("wrv_byalias")
                    vChange.Comment   = vRS("wrv_comment")
                    vRS.MoveNext
                Loop
            End If
            vRS.Close

            ' TODO: move this out of this method
            ' If this is the RecentChanges page, then force the presence of the
            ' <RecentChanges> element in the page.
            If vPage.Name = OPENWIKI_RCNAME Then
                vPage.Text = s(vPage.Text, "\<RecentChanges\>", "<RecentChangesLong>", True, True)
                If Not m(vPage.Text, "\<RecentChangesLong\>", True, True) Then
                    vPage.Text = vPage.Text & "<RecentChangesLong>"
                End If
            End If

            Call SetCachedPage(pPageName, pRevision, pIncludeText, pIncludeAllChangeRecords, vPage)
        End If

        Set GetPage = vPage
    End Function

    Function GetPageCount()
        vQuery = "SELECT COUNT(*) FROM openwiki_pages"
        vRS.Open vQuery, vConn, adOpenForwardOnly
        GetPageCount = CInt(vRS(0))
        vRS.Close
    End Function

    Function GetRevisionsCount()
        vQuery = "SELECT COUNT(*) FROM openwiki_revisions"
        vRS.Open vQuery, vConn, adOpenForwardOnly
        GetRevisionsCount = CInt(vRS(0))
        vRS.Close
    End Function

    Function ToXML(pXmlStr)
        ToXML = "<ow:wiki version='" & OPENWIKI_XMLVERSION & "' xmlns:ow='" & OPENWIKI_NAMESPACE & "' encoding='" & OPENWIKI_ENCODING & "' mode='" & gAction & "'>" _
              & "<ow:useragent>" & PCDATAEncode(Request.ServerVariables("HTTP_USER_AGENT")) & "</ow:useragent>" _
              & "<ow:location>" & PCDATAEncode(gServerRoot) & "</ow:location>" _
              & "<ow:scriptname>" & PCDATAEncode(gScriptName) & "</ow:scriptname>" _
              & "<ow:imagepath>" & PCDATAEncode(OPENWIKI_IMAGEPATH) & "</ow:imagepath>" _
              & "<ow:iconpath>" & PCDATAEncode(OPENWIKI_ICONPATH) & "</ow:iconpath>" _
              & "<ow:about>" & PCDATAEncode(gServerRoot & gScriptName & "?" & Request.ServerVariables("QUERY_STRING")) & "</ow:about>" _
              & "<ow:title>" & PCDATAEncode(OPENWIKI_TITLE) & "</ow:title>" _
              & "<ow:frontpage name='" & CDATAEncode(OPENWIKI_FRONTPAGE) & "' href='" & gScriptName & "?" & Server.URLEncode(OPENWIKI_FRONTPAGE) & "'>" & PCDATAEncode(PrettyWikiLink(OPENWIKI_FRONTPAGE)) & "</ow:frontpage>"
        If cEmbeddedMode = 0 Then
            If cAllowAttachments = 1 Then
                ToXML = ToXML & "<ow:allowattachments/>"
            End If
            If Request("redirect") <> "" Then
                ToXML = ToXML & "<ow:redirectedfrom name='" & CDATAEncode(URLDecode(Request("redirect"))) & "'>" & PCDATAEncode(PrettyWikiLink(URLDecode(Request("redirect")))) & "</ow:redirectedfrom>"
            End If
            ToXML = ToXML & getUserPreferences() & GetCookieTrail()
        End If
        ToXML = ToXML & pXmlStr & "</ow:wiki>"
    End Function

    Private Function isValidDocument(pText)
        On Error Resume Next
        Dim vXmlStr, vXmlDoc
        vXmlStr = "<ow:wiki xmlns:ow='x'>" & Wikify(pText) & "</ow:wiki>"
        If MSXML_VERSION = 4 Then
            Set vXmlDoc = Server.CreateObject("Msxml2.FreeThreadedDOMDocument.4.0")
        Else
            Set vXmlDoc = Server.CreateObject("Msxml2.FreeThreadedDOMDocument")
        End If
        vXmlDoc.async = False
        If vXmlDoc.loadXML(vXmlStr) Then
            isValidDocument = True
        Else
            isValidDocument = False
            Response.Write("<h1>Error occured</h1>")
            Response.Write("<b>Your input did not validate to well-formed and valid Wiki format.<br />")
            Response.Write("Please go back and correct. The XML output attempt was:</b><br /><br />")
            Response.Write("<pre>" & vbCRLF & Server.HTMLEncode(vXmlStr) & vbCRLF & "</pre>" & vbCRLF)
        End If
    End Function

    Function SavePage(pRevision, pMinorEdit, pComment, pText)
        Dim vRevision, vStatus, vHost, vUserAgent, vBy, vByAlias, vReplacedTS, vRevsDeleted

        pText = pText & ""
        If Not isValidDocument(pText) Then
            SavePage = False
            Response.End
        End If

        vHost = GetRemoteHost()
        vUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
        vBy = GetRemoteUser()
        If vBy = "" Then
            vBy = vHost
        End If
        vByAlias = GetRemoteAlias()

        Dim conn
        Set conn = Server.CreateObject("ADODB.Connection")
        conn.Open OPENWIKI_DB
        BeginTrans(conn)
        vQuery = "SELECT * FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_current = 1"
        vRS.Open vQuery, conn, adOpenKeyset, adLockOptimistic, adCmdText
        If vRS.EOF Then
            If Trim(pText) = "" Then
                RollbackTrans(conn)
                conn.Close()
                Set conn = Nothing
                SavePage = True
                Exit Function
            End If
            vRevision  = 1
            vStatus    = 1  ' new
        Elseif vRS("wrv_text") = pText Then
            RollbackTrans(conn)
            conn.Close()
            Set conn = Nothing
            SavePage = True
            Exit Function
        Else
            If (CInt(vRS("wrv_revision")) <> (pRevision - 1)) Then
                If ((vRS("wrv_by") <> vBy) Or (vRS("wrv_host") <> vHost) Or (vRS("wrv_agent") <> vUserAgent)) Then
                    RollbackTrans(conn)
                    conn.Close()
                    Set conn = Nothing
                    SavePage = False
                    Exit Function
                End If
            End If
            vRevision = CInt(vRS("wrv_revision")) + 1
            If ((vRS("wrv_by") = vBy) And (vRS("wrv_host") = vHost) And (vRS("wrv_agent") = vUserAgent)) Then
                vStatus = CInt(vRS("wrv_status"))
            Else
                vStatus = 2  ' updated
            End If
        End If

        If InStr(pText, "#DEPRECATED") = 1 Then
            vStatus = 3  ' deleted
        Elseif vStatus = 3 Then
            vStatus = 2  ' updated
        End If

        If vRS.EOF Then
            vQuery = "INSERT INTO openwiki_pages (wpg_name, wpg_lastminor, wpg_changes, wpg_lastmajor) VALUES " _
                   & "('" & Replace(gPage, "'", "''") & "'," & vRevision & " ,1 ," & vRevision & ")"
            conn.execute(vQuery)
        Else
            vQuery = "UPDATE openwiki_pages " _
                   & "SET wpg_changes = wpg_changes + 1" _
                   & ",   wpg_lastminor = " & vRevision
            If pMinorEdit = 0 Then
                vQuery = vQuery & ", wpg_lastmajor = " & vRevision
            End If
            vQuery = vQuery & " WHERE wpg_name = '" & Replace(gPage, "'", "''") & "'"
            conn.execute(vQuery)

            vQuery = "UPDATE openwiki_revisions SET wrv_current = 0 WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_current = 1"
            conn.execute(vQuery)
        End If
        vRS.Close

        vRS.Open "openwiki_revisions", conn, adOpenKeyset, adLockOptimistic, adCmdTable
        vRS.AddNew
        vRS("wrv_name")       = gPage
        vRS("wrv_revision")   = vRevision
        vRS("wrv_current")    = 1
        vRS("wrv_status")     = vStatus
        vRS("wrv_timestamp")  = Now()
        vRS("wrv_minoredit")  = pMinorEdit
        vRS("wrv_host")       = vHost
        vRS("wrv_agent")      = vUserAgent
        vRS("wrv_by")         = vBy
        vRS("wrv_byalias")    = vByAlias
        vRS("wrv_comment")    = pComment
        vRS("wrv_text")       = pText
        vRS.Update
        vRS.Close

        ' delete old revisions
        vQuery = "SELECT wrv_revision, wrv_timestamp FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' ORDER BY wrv_revision DESC"
        vRS.Open vQuery, conn, adOpenKeyset, adLockOptimistic, adCmdText
        If Not vRS.EOF Then
            ' this is the current revision
            vRS.MoveNext
            If Not vRS.EOF Then
                vReplacedTS = vRS("wrv_timestamp")
                ' keep at least one old revision
                vRS.MoveNext
                Do While Not vRS.EOF
                    ' check the timestamp of revision that replaced this revision
                    If vReplacedTS < (Now() - OPENWIKI_DAYSTOKEEP) Then
                        vQuery = "DELETE FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_revision <= " & CInt(vRS("wrv_revision"))
                        conn.execute(vQuery)
                        vRS.Close
                        vQuery = "SELECT COUNT(*) FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "'"
                        vRS.Open vQuery, conn, adOpenKeyset, adLockOptimistic, adCmdText
                        vQuery = "UPDATE openwiki_pages SET wpg_changes = " & CInt(vRS(0)) & " WHERE wpg_name = '" & Replace(gPage, "'", "''") & "'"
                        conn.execute(vQuery)
                        Exit Do
                    Else
                        vReplacedTS = vRS("wrv_timestamp")
                    End If
                    vRS.MoveNext
                Loop
            End If
        End If
        vRS.Close

        ' throw out the bath and the bathwater. TODO: keep the bath
        ClearDocumentCache(conn)

        CommitTrans(conn)
        conn.Close()

        Set conn = Nothing

        SavePage = True
    End Function


    ' returns the name of the file as you should save it
    ' pStatus : 0 = normal, 1 = hidden, 2 = deprecated
    Function SaveAttachmentMetaData(pFilename, pFilesize, pAddLink, pHidden, pComment)
        Dim vHost, vUserAgent, vBy, vByAlias, vPageRevision, vFileRevision, vFilename, vPos

        pFilename = Replace(pFilename, " ", "_")

        If pHidden = "" Then
            pHidden = 0
        End If

        vHost = GetRemoteHost()
        vUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
        vBy = GetRemoteUser()
        If vBy = "" Then
            vBy = vHost
        End If
        vByAlias = GetRemoteAlias()

        vQuery = "SELECT wpg_lastminor FROM openwiki_pages WHERE wpg_name = '" & Replace(gPage, "'", "''") & "'"
        vRS.Open vQuery, vConn, adOpenForwardOnly
        If vRS.EOF Then
            vPageRevision = 1 ' page doesn't exist yet
        Else
            vPageRevision = CInt(vRS(0))
        End If
        vRS.Close
        vQuery = "SELECT MAX(att_revision) FROM openwiki_attachments WHERE att_wrv_name = '" & Replace(gPage, "'", "''") & "' AND att_name = '" & Replace(pFilename, "'", "''") & "'"
        vRS.Open vQuery, vConn, adOpenForwardOnly
        If IsNull(vRS(0)) Then
            vFileRevision = 1
        Else
            vFileRevision = CInt(vRS(0)) + 1
        End If
        vRS.Close

        vPos = InStrRev(pFilename, ".")
        If vPos > 0 Then
            vFilename = Left(pFilename, vPos - 1) & "-" & vFileRevision & Mid(pFilename, vPos)
        Else
            vFilename = pFilename & "-" & vFileRevision
        End If
        vFilename = SafeFileName(vFilename)

        BeginTrans(vConn)
        vRS.Open "openwiki_attachments", vConn, adOpenKeyset, adLockOptimistic, adCmdTable
        vRS.AddNew
        vRS("att_wrv_name")     = gPage
        vRS("att_wrv_revision") = vPageRevision
        vRS("att_name")         = pFilename
        vRS("att_revision")     = vFileRevision
        vRS("att_hidden")       = pHidden
        vRS("att_deprecated")   = 0
        vRS("att_filename")     = vFilename
        vRS("att_timestamp")    = Now()
        vRS("att_filesize")     = pFilesize
        vRS("att_host")         = vHost
        vRS("att_agent")        = vUserAgent
        vRS("att_by")           = vBy
        vRS("att_byalias")      = vByAlias
        vRS("att_comment")      = pComment
        vRS.Update
        vRS.Close

        Call SaveAttachmentLog(vConn, pFilename, vFileRevision, "uploaded")

        Call ClearDocumentCache(vConn)
        'Call ClearDocumentCache2(vConn, gPage)

        If pAddLink <> "" Then
            If OPENWIKI_DB_SYNTAX = DB_MYSQL Then
                vQuery = "UPDATE openwiki_revisions SET wrv_text = CONCAT(wrv_text, '" & vbCRLF & vbCRLF & "  * " & Replace(pFilename, "'", "''") & "') WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_current = 1"
                vConn.execute(vQuery)
            Else
                vQuery = "SELECT wrv_text FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_current = 1"
                vRS.Open vQuery, vConn, adOpenKeyset, adLockOptimistic, adCmdText
                If Not vRS.EOF Then
                    vRS("wrv_text") = vRS("wrv_text") & vbCRLF & vbCRLF & "  * " & pFilename
                    vRS.Update
                End If
                vRS.Close
            End If
        End If

        CommitTrans(vConn)

        SaveAttachmentMetaData = vFilename
    End Function


    Function HideAttachmentMetaData(pName, pRevision, pHide)
        BeginTrans(vConn)
        vConn.Execute "UPDATE openwiki_attachments SET att_hidden = " & pHide & " WHERE att_wrv_name = '" & Replace(gPage, "'", "''") & "' AND att_name = '" & Replace(pName, "'", "''") & "' AND att_revision = " & pRevision
        If pHide = 1 Then
            Call SaveAttachmentLog(vConn, pName, pRevision, "hidden")
        Else
            Call SaveAttachmentLog(vConn, pName, pRevision, "made visible")
        End If
        Call ClearDocumentCache(vConn)
        'Call ClearDocumentCache2(vConn, gPage)
        CommitTrans(vConn)
    End Function


    Function TrashAttachmentMetaData(pName, pRevision, pTrash)
        BeginTrans(vConn)
        vConn.Execute "UPDATE openwiki_attachments SET att_deprecated = " & pTrash & " WHERE att_wrv_name = '" & Replace(gPage, "'", "''") & "' AND att_name = '" & Replace(pName, "'", "''") & "'"
        If pTrash = 1 Then
            Call SaveAttachmentLog(vConn, pName, pRevision, "deprecated")
        Else
            Call SaveAttachmentLog(vConn, pName, pRevision, "restored")
        End If
        Call ClearDocumentCache(vConn)
        'Call ClearDocumentCache2(vConn, gPage)
        CommitTrans(vConn)
    End Function


    Sub SaveAttachmentLog(pConn, pName, pFileRevision, pAction)
        Dim vHost, vUserAgent, vBy, vByAlias
        Dim pPagename, pPageRevision

        vHost = GetRemoteHost()
        vUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
        vBy = GetRemoteUser()
        If vBy = "" Then
            vBy = vHost
        End If
        vByAlias = GetRemoteAlias()

        vQuery = "SELECT att_wrv_name, att_wrv_revision FROM openwiki_attachments WHERE att_wrv_name = '" & Replace(gPage, "'", "''") & "' AND att_name = '" & Replace(pName, "'", "''") & "' AND att_revision = " & pFileRevision
        vRS.Open vQuery, pConn, adOpenForwardOnly
        If vRS.EOF Then
            vRS.Close
            Exit Sub
        End If
        pPagename = vRS("att_wrv_name")
        pPageRevision = vRS("att_wrv_revision")
        vRS.Close

        vQuery = "SELECT wrv_timestamp FROM openwiki_revisions WHERE wrv_name = '" & Replace(pPagename, "'", "''") & "' AND wrv_revision = " & pPageRevision
        vRS.Open vQuery, pConn, adOpenKeyset, adLockOptimistic, adCmdText
        If vRS.EOF Then
            vRS.Close
            Exit Sub
        End If
        vRS("wrv_timestamp") = Now()
        vRS.Update
        vRS.Close

        vRS.Open "openwiki_attachments_log", pConn, adOpenKeyset, adLockOptimistic, adCmdTable
        vRS.AddNew
        vRS("ath_wrv_name")     = pPagename
        vRS("ath_wrv_revision") = pPageRevision
        vRS("ath_name")         = pName
        vRS("ath_revision")     = pFileRevision
        vRS("ath_timestamp")    = Now()
        vRS("ath_agent")        = vUserAgent
        vRS("ath_by")           = vBy
        vRS("ath_byalias")      = vByAlias
        vRS("ath_action")       = pAction
        vRS.Update
        vRS.Close
    End Sub


    ' Convert the filename to a filename with an extension that is safe
    ' to be served by the webserver.
    Function SafeFileName(pFilename)
        Dim vPos, vExtension
        SafeFileName = pFilename
        vPos = InStrRev(pFilename, ".")
        If vPos > 0 Then
            vExtension = Mid(pFilename, vPos + 1)
            If gNotAcceptedExtensions = "" Then
                ' accept nothing, except the ones enumerated in gDocExtensions
                If Not InStr("|" & gDocExtensions & "|", "|" & vExtension & "|") > 0 Then
                    SafeFileName = SafeFilename & ".safe"
                End If
            Else
                ' accept everything, except the ones enumerated in gNotAcceptedExtensions
                If InStr("|" & gNotAcceptedExtensions & "|", "|" & vExtension & "|") > 0 Then
                    SafeFileName = SafeFilename & ".safe"
                End If
            End If
        End If
    End Function


    Sub GetAttachments(pPage, pRevision, pIncludeAllChangeRecords)
        Dim vAttachment, vMaxRevision
        If pIncludeAllChangeRecords Then
            ' show all file revisions
            vQuery = "SELECT att_name, att_revision, att_hidden, att_deprecated, att_filename, att_timestamp, att_filesize, att_by, att_byalias, att_comment" _
            & " FROM openwiki_attachments" _
            & " WHERE att_wrv_name = '" & Replace(pPage.Name, "'", "''") & "'" _
            & " AND   att_name = '" & Replace(Request("file"), "'", "''") & "'" _
            & " ORDER BY att_revision DESC"
        Else
            ' show last file revision relative to page revision
            vQuery = "SELECT MAX(att_wrv_revision) FROM openwiki_attachments WHERE att_wrv_name = '" & Replace(pPage.Name, "'", "''") & "'"
            If pRevision > 0 Then
                vQuery = vQuery & " AND att_wrv_revision <= " & pRevision
            End If
            vRS.Open vQuery, vConn, adOpenForwardOnly
            If IsNull(vRS(0)) Then
                vMaxRevision = 0
            Else
                vMaxRevision = CInt(vRS(0))
            End If
            vRS.Close
            vQuery = "SELECT att_name, att_revision, att_hidden, att_deprecated, att_filename, att_timestamp, att_filesize, att_by, att_byalias, att_comment" _
            & " FROM openwiki_attachments" _
            & " WHERE att_wrv_name = '" & Replace(pPage.Name, "'", "''") & "'" _
            & " AND   att_wrv_revision <= " & vMaxRevision _
            & " ORDER BY att_name ASC, att_revision DESC"
        End If
        vRS.Open vQuery, vConn, adOpenForwardOnly
        Do While Not vRS.EOF
            Set vAttachment = New Attachment
            vAttachment.Name       = vRS("att_name")
            vAttachment.Revision   = CInt(vRS("att_revision"))
            vAttachment.Hidden     = CInt(vRS("att_hidden"))
            vAttachment.Deprecated = CInt(vRS("att_deprecated"))
            vAttachment.Filename   = vRS("att_filename")
            vAttachment.Timestamp  = vRS("att_timestamp")
            vAttachment.Filesize   = CLng(vRS("att_filesize"))
            vAttachment.By         = vRS("att_by")
            vAttachment.ByAlias    = vRS("att_byalias")
            vAttachment.Comment    = vRS("att_comment")
            Call pPage.AddAttachment(vAttachment, Not pIncludeAllChangeRecords)
            vRS.MoveNext
        Loop
        vRS.Close
        pPage.AttachmentsLoaded = True
    End Sub


    ' pFilter --> 0=All, 1=NoMinorEdit, 2=OnlyMinorEdit
    Function TitleSearch(pPattern, pDays, pFilter, pOrderBy, pIncludeAttachmentChanges)
        Dim vTitle, vRegEx, vList, vPage, vChange, vCurPage, vAttachmentChange
        Set vList = New Vector
        Set vRegEx = New RegExp
        vRegEx.IgnoreCase = True
        vRegEx.Global = True
        vRegEx.Pattern = EscapePattern(pPattern)
        vQuery = "SELECT wpg_name, wpg_changes, wrv_revision, wrv_status, wrv_timestamp, wrv_minoredit, wrv_by, wrv_byalias, wrv_comment "
        If cAllowAttachments AND pIncludeAttachmentChanges Then
            vQuery = vQuery & ", ath_name, ath_revision, ath_timestamp, ath_by, ath_byalias, ath_action "
            If OPENWIKI_DB_SYNTAX = DB_ORACLE Then
                vQuery = vQuery & " FROM   openwiki_pages, openwiki_revisions, openwiki_attachments_log " _
                                & " WHERE  wpg_name = wrv_name " _
                                & " AND    wrv_name = ath_wrv_name (+) " _
                                & " AND    wrv_revision = ath_wrv_revision (+)"
            Else
                vQuery = vQuery & " FROM (openwiki_pages LEFT JOIN openwiki_revisions ON openwiki_pages.wpg_name = openwiki_revisions.wrv_name) LEFT JOIN openwiki_attachments_log ON (openwiki_revisions.wrv_name = openwiki_attachments_log.ath_wrv_name) AND (openwiki_revisions.wrv_revision = openwiki_attachments_log.ath_wrv_revision) WHERE 1 = 1 "
            End If
        Else
            vQuery = vQuery & "FROM openwiki_pages, openwiki_revisions " _
                            & "WHERE wrv_name = wpg_name "
        End If
        If pDays > 0 Then
            ' is there a database independent way to test the current date?
            'vQuery = vQuery & " AND wpg_timestamp >
        End If
        If pFilter = 0 Then
            vQuery = vQuery & " AND wpg_lastminor = wrv_revision"
        Elseif pFilter = 1 Then
            vQuery = vQuery & " AND wpg_lastmajor = wrv_revision"
        Elseif pFilter = 2 Then
            vQuery = vQuery & " AND wpg_lastminor = wrv_revision AND wrv_minoredit = 1"
        End If
        If pOrderBy = 1 Then
            vQuery = vQuery & " ORDER BY wrv_timestamp DESC"
        Elseif pOrderBy = 2 Then
            vQuery = vQuery & " ORDER BY wrv_timestamp"
        Else
            vQuery = vQuery & " ORDER BY wpg_name"
        End If
        If cAllowAttachments AND pIncludeAttachmentChanges Then
            vQuery = vQuery & ", ath_timestamp DESC"
        End If
        vRS.Open vQuery, vConn, adOpenForwardOnly
        Do While Not vRS.EOF
            If vRegEx.Test(vRS("wpg_name")) Then
                If vCurPage <> vRS("wpg_name") Then
                    vCurPage = vRS("wpg_name")
                    Set vPage = New WikiPage
                    vPage.Name = vRS("wpg_name")
                    vPage.Changes = CInt(vRS("wpg_changes"))
                    Set vChange = vPage.AddChange
                    vChange.Revision  = CInt(vRS("wrv_revision"))
                    vChange.Status    = CInt(vRS("wrv_status"))
                    vChange.MinorEdit = CInt(vRS("wrv_minoredit"))
                    vChange.Timestamp = vRS("wrv_timestamp")
                    vChange.By        = vRS("wrv_by")
                    vChange.ByAlias   = vRS("wrv_byalias")
                    vChange.Comment   = vRS("wrv_comment")
                    vList.Push(vPage)
                End If
                If cAllowAttachments AND pIncludeAttachmentChanges Then
                    If (vRS("ath_name") <> "") And (vRS("ath_timestamp") > DateAdd("h", -24, Now())) Then
                        Set vAttachmentChange = New AttachmentChange
                        vAttachmentChange.Name = vRS("ath_name")
                        vAttachmentChange.Revision = CInt(vRS("ath_revision"))
                        vAttachmentChange.Timestamp = vRS("ath_timestamp")
                        vAttachmentChange.By = vRS("ath_by")
                        vAttachmentChange.ByAlias = vRS("ath_byalias")
                        vAttachmentChange.Action = vRS("ath_action")
                        vChange.AddAttachmentChange(vAttachmentChange)
                    End If
                End If
            End If
            vRS.MoveNext
        Loop
        vRS.Close
        Set vRegEx = Nothing
        Set TitleSearch = vList
    End Function


    Function FullSearch(pPattern, pIncludeTitles)
        Dim vTitle, vRegEx, vRegEx2, vList, vPage, vChange, vFound
        pPattern = EscapePattern(pPattern)
        Set vList = New Vector
        Set vRegEx = New RegExp
        vRegEx.IgnoreCase = True
        vRegEx.Global = True
        If Request("fromtitle") = "true" Then
            vRegEx.Pattern = Replace(pPattern, "_", " ")
        Else
            vRegEx.Pattern = pPattern
        End If
        If pIncludeTitles Then
            Set vRegEx2 = New RegExp
            vRegEx2.IgnoreCase = True
            vRegEx2.Global = True
            vRegEx2.Pattern = pPattern
        End If
        vQuery = "SELECT * FROM openwiki_pages, openwiki_revisions WHERE wrv_name = wpg_name AND wrv_current = 1 AND wrv_text IS NOT NULL ORDER BY wpg_name"
        vRS.Open vQuery, vConn, adOpenForwardOnly
        Do While Not vRS.EOF
            vFound = False
            If pIncludeTitles Then
                If vRegEx2.Test(vRS("wpg_name")) Then
                    vFound = True
                End If
            End If
            If Not vFound Then
                If vRegEx.Test(vRS("wrv_text")) Then
                    vFound = True
                End If
            End If
            If vFound Then
                Set vPage = New WikiPage
                vPage.Name = vRS("wpg_name")
                vPage.Changes = CInt(vRS("wpg_changes"))
                Set vChange = vPage.AddChange
                vChange.Revision  = CInt(vRS("wrv_revision"))
                vChange.Status    = CInt(vRS("wrv_status"))
                vChange.MinorEdit = CInt(vRS("wrv_minoredit"))
                vChange.Timestamp = vRS("wrv_timestamp")
                vChange.By        = vRS("wrv_by")
                vChange.ByAlias   = vRS("wrv_byalias")
                vChange.Comment   = vRS("wrv_comment")
                vList.Push(vPage)
            End If
            vRS.MoveNext
        Loop
        vRS.Close
        Set vRegEx = Nothing
        Set vRegEx2 = Nothing
        Set FullSearch = vList
    End Function


    Function GetPreviousRevision(pDiffType, pDiffTo)
        Dim vBy, vHost, vAgent
        GetPreviousRevision = 0
        If pDiffTo <= 0 Then
            pDiffTo = 99999999
        End If
        vQuery = "SELECT wrv_revision, wrv_minoredit, wrv_by, wrv_host, wrv_agent FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_revision <= " & pDiffTo
        vQuery = vQuery & " ORDER BY wrv_revision DESC"
        vRS.Open vQuery, vConn, adOpenForwardOnly
        If Not vRS.EOF Then
            vBy    = vRS("wrv_by")
            vHost  = vRS("wrv_host")
            vAgent = vRS("wrv_agent")
        End If
        Do While Not vRS.EOF
            GetPreviousRevision = CInt(vRS("wrv_revision"))
            If pDiffType = 0 Then
                ' previous major
                If CInt(vRS("wrv_minoredit")) = 0 Then
                    vRS.MoveNext
                    If Not vRS.EOF Then
                        GetPreviousRevision = CInt(vRS("wrv_revision"))
                    End If
                    Exit Do
                End If
            Elseif pDiffType = 1 Then
                ' previous minor
                vRS.MoveNext
                If Not vRS.EOF Then
                    GetPreviousRevision = CInt(vRS("wrv_revision"))
                End If
                Exit Do
            Else
                ' previous author
                If vRS("wrv_by") <> vBy Or vRS("wrv_host") <> vHost Or vRS("wrv_agent") <> vAgent Then
                    Exit Do
                End If
            End If
            vRS.MoveNext
        Loop
        vRS.Close
    End Function


    Function InterWiki()
        Dim vTemp
        vQuery = "SELECT wik_name, wik_url FROM openwiki_interwikis ORDER BY wik_name"
        vRS.Open vQuery, vConn, adOpenForwardOnly
        Do While Not vRS.EOF
            Dim val
            vTemp = vTemp & "<ow:interlink href='" & CDATAEncode(vRS("wik_url")) & "'>" & PCDATAEncode(vRS("wik_name")) & "</ow:interlink>"
            vRS.MoveNext
        Loop
        vRS.Close
        InterWiki = "<ow:interlinks>" & vTemp & "</ow:interlinks>"
    End Function


    Function GetInterWiki(pName)
        If OPENWIKI_DB <> "" Then
            If pName = "This" Then
                GetInterWiki = gScriptName & "?p="
            Else
                vQuery = "SELECT wik_url FROM openwiki_interwikis WHERE wik_name = '" & Replace(pName, "'", "''") & "'"
                vRS.Open vQuery, vConn, adOpenForwardOnly
                If Not vRS.EOF Then
                    GetInterWiki = vRS("wik_url")
                End If
                vRS.Close
            End If
        End If
    End Function


    Function GetRSSFromCache(pURL, pRefreshRate, pFreshlyFromRemoteSite, pRetryLater)
        Dim conn, vRS
        Dim vLast, vNext, vRefreshRate
        Set conn = Server.CreateObject("ADODB.Connection")
        conn.Open OPENWIKI_DB
        vQuery = "SELECT rss_last, rss_next, rss_refreshrate, rss_cache FROM openwiki_rss WHERE rss_url = '" & Replace(pURL, "'", "''") & "'"
        Set vRS = Server.CreateObject("ADODB.Recordset")
        vRS.Open vQuery, conn, adOpenKeyset, adLockOptimistic, adCmdText
        If vRS.EOF Then
            GetRSSFromCache = "notexists"
        Else
            vLast        = vRS("rss_last")
            vNext        = vRS("rss_next")
            vRefreshRate = CInt(vRS("rss_refreshrate"))
            If vRefreshRate <> pRefreshRate Then
                vNext = DateAdd("n", pRefreshRate, vLast)
                vRS("rss_next")        = vNext
                vRS("rss_refreshrate") = pRefreshRate
                vRS.Update
            Elseif pRetryLater Then
                ' retry a minute from now
                vNext = DateAdd("n", 1, Now())
                vRS("rss_next") = vNext
                vRS.Update
            End If

            If pFreshlyFromRemoteSite Or (DateDiff("n", vNext, Now()) < 0) Then
                GetRSSFromCache = "<ow:feed href='" & Replace(pURL, "&", "&amp;") & "' "
                If pFreshlyFromRemoteSite Then
                    GetRSSFromCache = GetRSSFromCache & "fresh='true' "
                Else
                    GetRSSFromCache = GetRSSFromCache & "fresh='false' "
                End If
                GetRSSFromCache = GetRSSFromCache & "last='" & FormatDateISO8601(vLast) & "' "
                GetRSSFromCache = GetRSSFromCache & "next='" & FormatDateISO8601(vNext) & "' "
                GetRSSFromCache = GetRSSFromCache & "refreshrate='" & pRefreshRate & "'>"
                GetRSSFromCache = GetRSSFromCache & vRS("rss_cache")
                GetRSSFromCache = GetRSSFromCache & "</ow:feed>"
            End If

        End If
        vRS.Close
        conn.Close
        Set vRS = Nothing
        Set conn = Nothing
    End Function

    Sub SaveRSSToCache(pURL, pRefreshRate, pCache)
        Dim conn, vRS
        Set conn = Server.CreateObject("ADODB.Connection")
        conn.Open OPENWIKI_DB
        vQuery = "SELECT * FROM openwiki_rss WHERE rss_url = '" & Replace(pURL, "'", "''") & "'"
        Set vRS = Server.CreateObject("ADODB.Recordset")
        vRS.Open vQuery, conn, adOpenKeyset, adLockOptimistic, adCmdText
        If vRS.EOF Then
            vRS.Close
            vRS.Open "openwiki_rss", conn, adOpenKeyset, adLockOptimistic, adCmdTable
            vRS.AddNew
            vRS("rss_url") = pURL
        End If
        vRS("rss_last")        = Now()
        If pCache = "" Then
            vRS("rss_next") = DateAdd("n", 30, Now())   ' 30 minutes from now
        Else
            vRS("rss_next") = DateAdd("n", pRefreshRate, Now())
        End If
        vRS("rss_refreshrate") = pRefreshRate
        vRS("rss_cache")       = pCache
        vRS.Update
        vRS.Close
        conn.Close
        Set vRS = Nothing
        Set conn = Nothing
    End Sub

    Sub Aggregate(pURL, pXmlDoc)
        Dim conn, vRS
        Dim vRoot, vItems, vItem
        Dim vXmlIsland, vAgXmlIsland, vNow, i
        Dim vRssLink, vRdfResource, vRdfTimestamp, vDcDate

        On Error Resume Next
        'Response.Write("<p />Aggregating " & pURL & "<br />")

        Set vRoot = pXmlDoc.documentElement

        If vRoot.NodeName = "rss" Then
            Set vItems = vRoot.selectNodes("channel/item")
        Elseif vRoot.getAttribute("xmlns") = "http://my.netscape.com/rdf/simple/0.9/" Then
            Set vItems = vRoot.selectNodes("item")
        Elseif vRoot.getAttribute("xmlns") = "http://purl.org/rss/1.0/" Then
            Set vItems = vRoot.selectNodes("item")
        Else
            Exit Sub
        End If

        vNow = Now()
        i = 0

        ' TODO: find workaround for bug in MSXML v4
        If Not vRoot.selectSingleNode("channel/wiki:interwiki") Is Nothing Then
            vAgXmlIsland = "<ag:source><rdf:Description wiki:interwiki=""" & vRoot.selectSingleNode("channel/wiki:interwiki").text & """><rdf:value>" & PCDATAEncode(vRoot.selectSingleNode("channel/title").text) & "</rdf:value></rdf:Description></ag:source>"
        Else
            vAgXmlIsland = "<ag:source>" & PCDATAEncode(vRoot.selectSingleNode("channel/title").text) & "</ag:source>"
        End If
        vAgXmlIsland = vAgXmlIsland & "<ag:sourceURL>" & PCDATAEncode(vRoot.selectSingleNode("channel/link").text) & "</ag:sourceURL>"

        Set conn = Server.CreateObject("ADODB.Connection")
        conn.Open OPENWIKI_DB
        Set vRS = Server.CreateObject("ADODB.Recordset")

        ' walk trough all item elements and store them in the openwiki_rss_aggregations table
        For Each vItem In vItems
            vRssLink = vItem.selectSingleNode("link").text

            vRdfResource = vItem.getAttribute("rdf:about")
            If IsNull(vRdfResource) Then
                vRdfResource = vRssLink
            End If

            If vItem.selectSingleNode("ag:timestamp") Is Nothing Then
                vRdfTimestamp = DateAdd("s", i, vNow)
            Else
                vRdfTimestamp = vItem.selectSingleNode("ag:timestamp").text
                Call s(vRdfTimestamp, gTimestampPattern, "&ToDateTime($1,$2,$3,$4,$5,$6,$7,$8,$9)", False, False)
                If DateDiff("d", vNow, sReturn) > 1 Then
                    ' we cannot take this date serious, it's too far in the future
                    vRdfTimestamp = DateAdd("s", i, vNow)
                Else
                    vRdfTimestamp = sReturn
                    vAgXmlIsland = vItem.selectSingleNode("ag:source").xml & vItem.selectSingleNode("ag:sourceURL").xml
                End If
            End If
            i = i - 1

            vXmlIsland = "<title>" & PCDATAEncode(vItem.selectSingleNode("title").text) & "</title><link>" & PCDATAEncode(vItem.selectSingleNode("link").text) & "</link>"
            If Not vItem.selectSingleNode("description") Is Nothing Then
                vXmlIsland = vXmlIsland & "<description>" & PCDATAEncode(vItem.selectSingleNode("description").text) & "</description>"
            End If
            If Not vItem.selectSingleNode("dc:creator") Is Nothing Then
                vXmlIsland = vXmlIsland & vItem.selectSingleNode("dc:creator").xml
            End If
            If Not vItem.selectSingleNode("dc:contributor") Is Nothing Then
                vXmlIsland = vXmlIsland & vItem.selectSingleNode("dc:contributor").xml
            End If
            If vItem.selectSingleNode("dc:date") Is Nothing Then
                vDcDate = ""
            Else
                vDcDate = vItem.selectSingleNode("dc:date").text
                vXmlIsland = vXmlIsland & "<dc:date>" & vItem.selectSingleNode("dc:date").text & "</dc:date>"
            End If
            If Not vItem.selectSingleNode("wiki:version") Is Nothing Then
                vXmlIsland = vXmlIsland & "<wiki:version>" & vItem.selectSingleNode("wiki:version").text & "</wiki:version>"
            End If
            If Not vItem.selectSingleNode("wiki:status") Is Nothing Then
                vXmlIsland = vXmlIsland & "<wiki:status>" & vItem.selectSingleNode("wiki:status").text & "</wiki:status>"
            End If
            If Not vItem.selectSingleNode("wiki:importance") Is Nothing Then
                vXmlIsland = vXmlIsland & "<wiki:importance>" & vItem.selectSingleNode("wiki:importance").text & "</wiki:importance>"
            End If
            If Not vItem.selectSingleNode("wiki:diff") Is Nothing Then
                vXmlIsland = vXmlIsland & vItem.selectSingleNode("wiki:diff").xml
            End If
            If Not vItem.selectSingleNode("wiki:history") Is Nothing Then
                vXmlIsland = vXmlIsland & vItem.selectSingleNode("wiki:history").xml
            End If
            vXmlIsland = vXmlIsland & vAgXmlIsland & "<ag:timestamp>" & FormatDateISO8601(vRdfTimestamp) & "</ag:timestamp>"

            vXmlIsland = "<item rdf:about='" & PCDATAEncode(vRdfResource) & "'>" & vXmlIsland & "</item>"

            ' TODO: erm... this is actually inefficient.. use better ADO techniques
            vQuery = "SELECT * FROM openwiki_rss_aggregations WHERE agr_feed='" & Replace(pURL, "'", "''") & "' AND agr_rsslink = '" & Replace(vRssLink, "'", "''") & "'"
            vRS.Open vQuery, conn, adOpenKeyset, adLockOptimistic, adCmdText
            If vRS.EOF Then
                vRS.Close
                vRS.Open "openwiki_rss_aggregations", conn, adOpenKeyset, adLockOptimistic, adCmdTable
                vRS.AddNew
                vRS("agr_feed")      = pURL
                vRS("agr_resource")  = vRdfResource
                vRS("agr_rsslink")   = vRssLink
                vRS("agr_timestamp") = vRdfTimestamp
                vRS("agr_dcdate")    = vDcDate
                vRS("agr_xmlisland") = vXmlIsland
                vRS.Update
            Elseif vRS("agr_dcdate") <> vDcDate Then
                vRS("agr_resource")  = vRdfResource
                vRS("agr_timestamp") = vRdfTimestamp
                vRS("agr_dcdate")    = vDcDate
                vRS("agr_xmlisland") = vXmlIsland
                vRS.Update
            End If
            vRS.Close
        Next

        conn.Close
        Set vRS = Nothing
        Set conn = Nothing

        'Response.Write("<p />Done aggregating " & pURL & "<br />")
    End Sub

    Function GetAggregation(pURLs)
        Dim vRdfSeq, vItems, vTemp, i
        vQuery = ""
        Do While Not pURLs.IsEmpty
            vQuery = vQuery & "'" & Replace(pURLs.Pop(), "'", "''") & "'"
            If pURLs.Count >  0 Then
                vQuery = vQuery & ","
            End If
        Loop
        vQuery = "SELECT * FROM openwiki_rss_aggregations WHERE agr_feed IN (" & vQuery & ") ORDER BY agr_timestamp DESC"
        vRS.Open vQuery, vConn, adOpenForwardOnly
        i = 0
        If OPENWIKI_MAXNROFAGGR <= 0 Then
            OPENWIKI_MAXNROFAGGR = 100
        End If
        Do While Not vRS.EOF
            i = i + 1
            If i > OPENWIKI_MAXNROFAGGR Then
                Exit Do
            End If
            vTemp = CDATAEncode(vRS("agr_resource"))
            vRdfSeq = vRdfSeq & "<rdf:li rdf:resource='" & vTemp & "'/>"
            vItems = vItems & vRS("agr_xmlisland")
            vRS.MoveNext
        Loop
        vRS.Close
        GetAggregation = "<?xml version='1.0' encoding='ISO-8859-1'?>" & vbCRLF _
                      & "<!-- All Your Wiki Are Belong To Us -->" & vbCRLF _
                      & "<rdf:RDF xmlns='http://purl.org/rss/1.0/' xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:wiki='http://purl.org/rss/1.0/modules/wiki/' xmlns:ag='http://purl.org/rss/1.0/modules/aggregation/'>" _
                      & "<channel rdf:about='" & CDATAEncode(gServerRoot & gScriptName & "?p=" & gPage & "&a=rss") & "'>" _
                      & "<title>" & PCDATAEncode(OPENWIKI_TITLE & " -- " & PrettyWikiLink(gPage)) & "</title>" _
                      & "<link>" & PCDATAEncode(gServerRoot & gScriptName & "?" & gPage) & "</link>" _
                      & "<description>" & PCDATAEncode(OPENWIKI_TITLE & " -- " & PrettyWikiLink(gPage)) & "</description>" _
                      & "<image rdf:about='" & CDATAEncode(gServerRoot & "ow/images/aggregator.gif") & "'/>" _
                      & "<items><rdf:Seq>" _
                      & vRdfSeq _
                      & "</rdf:Seq></items>" _
                      & "</channel>" _
                      & "<image rdf:about='" & CDATAEncode(gServerRoot & "ow/images/aggregator.gif") & "'>" _
                      & "<title>" & PCDATAEncode(OPENWIKI_TITLE) & "</title>" _
                      & "<link>" & CDATAEncode(gServerRoot & gScriptName & "?p=" & gPage) & "</link>" _
                      & "<url>" & PCDATAEncode(gServerRoot & "ow/images/logo_aggregator.gif") & "</url>" _
                      & "</image>" _
                      & vItems _
                      & "</rdf:RDF>"
    End Function


    Private Function CreateDocKey(pSubKey)
        CreateDocKey = pSubKey & gFS _
                     & gCookieHash & gFS _
                     & gRevision & gFS _
                     & Request.Cookies(gCookieHash & "?up")("pwl") & gFS _
                     & Request.Cookies(gCookieHash & "?up")("new") & gFS _
                     & Request.Cookies(gCookieHash & "?up")("emo")
        CreateDocKey = Hash(CreateDocKey)
    End Function

    Function GetDocumentCache(pSubKey)
        vQuery = "SELECT chc_xmlisland FROM openwiki_cache WHERE chc_name = '" & Replace(gPage, "'", "''") & "' AND chc_hash = " & CreateDocKey(pSubKey)
        vRS.Open vQuery, vConn, adOpenForwardOnly
        If vRS.EOF Then
            GetDocumentCache = ""
        Else
            GetDocumentCache = vRS("chc_xmlisland")
        End If
        vRS.Close
    End Function

    Sub SetDocumentCache(pSubKey, pXmlStr)
        Dim vKey
        vKey = CreateDocKey(pSubKey)
        vQuery = "SELECT chc_xmlisland FROM openwiki_cache WHERE chc_name = '" & Replace(gPage, "'", "''") & "' AND chc_hash = " & vKey
        vRS.Open vQuery, vConn, adOpenKeyset, adLockOptimistic, adCmdText
        If vRS.EOF Then
            vRS.Close
            vRS.Open "openwiki_cache", vConn, adOpenKeyset, adLockOptimistic, adCmdTable
            vRS.AddNew
            vRS("chc_name") = gPage
            vRS("chc_hash") = vKey
        End If
        vRS("chc_xmlisland") = pXmlStr
        vRS.Update
        vRS.Close
    End Sub

    Sub ClearDocumentCache(pConn)
        pConn.Execute "DELETE FROM openwiki_cache"
    End Sub

    Sub ClearDocumentCache2(pConn, pPagename)
        If pConn = "" Then
            Set pConn = vConn
        End if
        pConn.Execute "DELETE FROM openwiki_cache WHERE chc_name = '" & Replace(pPagename, "'", "''") & "'"
    End Sub

End Class


'______________________________________________________________________________________________________________
Function FormatDateISO8601(pTimestamp)
    Dim vTemp
    FormatDateISO8601 = Year(pTimestamp) & "-"
    vTemp = Month(pTimestamp)
    If vTemp < 10 Then
        FormatDateISO8601 = FormatDateISO8601 & "0"
    End If
    FormatDateISO8601 = FormatDateISO8601 & vTemp & "-"
    vTemp = Day(pTimestamp)
    If vTemp < 10 Then
        FormatDateISO8601 = FormatDateISO8601 & "0"
    End If
    FormatDateISO8601 = FormatDateISO8601 & vTemp & "T"
    vTemp = Hour(pTimestamp)
    If vTemp < 10 Then
        FormatDateISO8601 = FormatDateISO8601 & "0"
    End If
    FormatDateISO8601 = FormatDateISO8601 & vTemp & ":"
    vTemp = Minute(pTimestamp)
    If vTemp < 10 Then
        FormatDateISO8601 = FormatDateISO8601 & "0"
    End If
    FormatDateISO8601 = FormatDateISO8601 & vTemp & ":"
    vTemp = Second(pTimestamp)
    If vTemp < 10 Then
        FormatDateISO8601 = FormatDateISO8601 & "0"
    End If
    FormatDateISO8601 = FormatDateISO8601 & vTemp
    FormatDateISO8601 = FormatDateISO8601 & OPENWIKI_TIMEZONE
End Function

Sub ToDateTime(pYear, pMonth, pDay, pHour, pMinutes, pSeconds, pPlusMinTZ, pHourTZ, pMinutesTZ)
    sReturn = DateSerial(pYear, pMonth, pDay)
    If pPlusMinTZ = "-" Then
        sReturn = DateAdd("h", pHour + pHourTZ, sReturn)
        sReturn = DateAdd("n", pMinutes + pMinutesTZ, sReturn)
    Elseif pPlusMinTZ = "+" Then
        sReturn = DateAdd("h", pHour - pHourTZ, sReturn)
        sReturn = DateAdd("n", pMinutes - pMinutesTZ, sReturn)
    End If
    If pPlusMinTZ = "-" Or pPlusMinTZ = "+" Then
        ' it's in GMT, now move it to OPENWIKI_TIMEZONE
        If Left(OPENWIKI_TIMEZONE, 1) = "-" Then
            sReturn = DateAdd("h", -1 * Mid(OPENWIKI_TIMEZONE, 2, 2), sReturn)
            sReturn = DateAdd("n", -1 * Mid(OPENWIKI_TIMEZONE, 5, 2), sReturn)
        Else
            sReturn = DateAdd("h", Mid(OPENWIKI_TIMEZONE, 2, 2), sReturn)
            sReturn = DateAdd("n", Mid(OPENWIKI_TIMEZONE, 5, 2), sReturn)
        End If
    End If
End Sub

Function EscapePattern(pPattern)
    Dim vRegEx
    pPattern = Replace(pPattern, "''''''", "")
    Set vRegEx = New RegExp
    vRegEx.IgnoreCase = True
    vRegEx.Global = True
    vRegEx.Pattern = pPattern
    On Error Resume Next
    Err.Number = 0
    vRegEx.Test("x")
    If Err.Number <> 0 Then
        pPattern = Replace(pPattern, "\", "\\")
        pPattern = Replace(pPattern, "(", "\(")
        pPattern = Replace(pPattern, ")", "\)")
        pPattern = Replace(pPattern, "[", "\[")
        pPattern = Replace(pPattern, "+", "\+")
        pPattern = Replace(pPattern, "*", "\*")
        pPattern = Replace(pPattern, "?", "\?")
    End If

    'Response.Write("Pattern : " & pPattern & "<br />")
    EscapePattern = pPattern
End Function

%>