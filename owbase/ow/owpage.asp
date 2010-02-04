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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owpage.asp,v $
'    $Revision: 1.3 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

Class WikiPage
    Private vName, vAnchor, vText, vRevision, vChanges, vLastMinor, vLastMajor
    Private vChangesCollection
    Private vAttachmentCollection
    Private vAttachmentsLoaded

    Private Sub Class_Initialize()
        vText     = ""
        vRevision = 0
        vChanges  = 0
        Set vChangesCollection = New Vector
        vAttachmentsLoaded = False
    End Sub

    Private Sub Class_Terminate()
        ' You are the weakest link. Goodbye.
        Set vChangesCollection = Nothing
        Set vAttachmentCollection = Nothing
    End Sub

    Public Property Get Name()
        Name = vName
    End Property

    Public Property Let Name(pName)
        vName = pName
    End Property

    Public Property Get Anchor()
        Anchor = vAnchor
    End Property

    Public Property Let Anchor(pAnchor)
        vAnchor = pAnchor
    End Property

    Public Property Get Revision()
        Revision = vRevision
    End Property

    Public Property Let Revision(pRevision)
        vRevision = pRevision
    End Property

    Public Property Get Changes()
        Changes = vChanges
    End Property

    Public Property Let Changes(pChanges)
        vChanges = pChanges
    End Property

    Public Property Get LastMinor()
        LastMinor = vLastMinor
    End Property

    Public Property Let LastMinor(pLastMinor)
        vLastMinor = pLastMinor
    End Property

    Public Property Get LastMajor()
        LastMajor = vLastMajor
    End Property

    Public Property Let LastMajor(pLastMajor)
        vLastMajor = pLastMajor
    End Property

    Public Property Get Text()
        Text = vText
    End Property

    Public Property Let Text(pText)
        vText = pText
    End Property

    Public Function AddChange()
        Dim vChange
        Set vChange = New Change
        vChangesCollection.Push(vChange)
        Set AddChange = vChange
    End Function

    Public Function GetLastChange()
        Set GetLastChange = vChangesCollection.ElementAt(0)
    End Function

    Public Function Exists()
        If vChangesCollection.ElementAt(0).Timestamp = "" Then
            Exists = False
        Else
            Exists = True
        End If
    End Function

    Public Property Get AttachmentsLoaded()
        AttachmentsLoaded = vAttachmentsLoaded
    End Property

    Public Property Let AttachmentsLoaded(pAttachmentsLoaded)
        vAttachmentsLoaded = pAttachmentsLoaded
    End Property

    Public Sub AddAttachment(pAttachment, pStoreMaxRevOnly)
        Dim i, vCount, vAttachment
        If Not IsObject(vAttachmentCollection) Then
            Set vAttachmentCollection = New Vector
        End If
        If pStoreMaxRevOnly Then
            For i = 0 To vAttachmentCollection.Count - 1
                Set vAttachment = vAttachmentCollection.ElementAt(i)
                If vAttachment.Name = pAttachment.Name Then
                    If vAttachment.Revision < pAttachment.Revision Then
                        vAttachmentCollection.RemoveElementAt(i)
                        vAttachmentCollection.Push(pAttachment)
                    End If
                    Exit Sub
                End If
            Next
        End If
        vAttachmentCollection.Push(pAttachment)
    End Sub

    Public Function GetAttachment(pName)
        Dim i, vCount, vAttachment
        If IsObject(vAttachmentCollection) Then
            For i = 0 To vAttachmentCollection.Count - 1
                Set vAttachment = vAttachmentCollection.ElementAt(i)
                If vAttachment.Name = pName Then
                    Set GetAttachment = vAttachment
                    Exit Function
                End If
            Next
        End If
        Set GetAttachment = Nothing
    End Function

    Public Function GetAttachmentPattern()
        Dim i, vCount, vAttachment
        GetAttachmentPattern = ""
        If IsObject(vAttachmentCollection) Then
            For i = 0 To vAttachmentCollection.Count - 1
                Set vAttachment = vAttachmentCollection.ElementAt(i)
                If i > 0 Then
                    GetAttachmentPattern = GetAttachmentPattern & "|"
                End If
                GetAttachmentPattern = GetAttachmentPattern & Replace(vAttachment.Name, ".", "\.")
            Next
        End If
    End Function

    Public Function ToLinkXML(pText, pTemplate, pAddPath)
        Dim vLastChange, vTemp
        Set vLastChange = vChangesCollection.ElementAt(0)
        If vLastChange.Timestamp = "" Then
            ToLinkXML = "<ow:link name='" & CDATAEncode(vName) & "' "
            If vAnchor <> "" Then
                ToLinkXML = ToLinkXML & " anchor='" & CDATAEncode(vAnchor) & "'"
            End If
            ToLinkXML = ToLinkXML & " href='" & gScriptName & "?p=" & Server.URLEncode(vName)
            If cDirectEdit = 1 Or (cTemplateLinking = 1 And pTemplate <> "") Then
                ToLinkXML = ToLinkXML & "&amp;a=edit"
            End If
            If cTemplateLinking = 1 And pTemplate <> "" Then
                ToLinkXML = ToLinkXML & "&amp;template=" & pTemplate
            End If
            ToLinkXML = ToLinkXML & "'>"
        Else
            If gAction = "print" Then
                vTemp = gScriptName & "?p=" & Server.URLEncode(vName) & "&amp;a=print"
            Else
                vTemp = gScriptName & "?" & Server.URLEncode(vName)
            End If
            ToLinkXML = "<ow:link name='" & CDATAEncode(vName) & "'"
            If vAnchor <> "" Then
                ToLinkXML = ToLinkXML & " anchor='" & CDATAEncode(vAnchor) & "'"
            End If
            ToLinkXML = ToLinkXML & " href='" & vTemp & "' date='" & FormatDateISO8601(vLastChange.Timestamp) & "'>"
        End If

        ToLinkXML = ToLinkXML & PCDATAEncode(pText) & "</ow:link>"
    End Function

    Public Function ToXML(pIncludeText)
        Dim i, vCount
        ToXML = "<ow:page name='" & CDATAEncode(vName) & "' changes='" & vChanges & "'"
        If vLastMinor > 0 Then
            ToXML = ToXML & " lastminor='" & vLastMinor & "'"
        End If
        If vLastMajor > 0 Then
            ToXML = ToXML & " lastmajor='" & vLastMajor & "'"
        End If
        If vRevision > 0 Then
            ToXML = ToXML & " revision='" & vRevision & "'"
        End If
        ToXML = ToXML & ">"
        ToXML = ToXML & ToLinkXML(PrettyWikiLink(vName), "", True)
        If vChangesCollection.ElementAt(0).Revision > 0 Then
            vCount = vChangesCollection.Count - 1
            For i = 0 To vCount
                ToXML = ToXML & vChangesCollection.ElementAt(i).ToXML()
            Next
        End If
        If pIncludeText = 1 Or pIncludeText = 3 Then
            If cEmbeddedMode = 0 And Trim(vText) = "" Then
                ToXML = gNamespace.GetIndexSchemes().GetTemplates(OPENWIKI_TEMPLATES) _
                      & ToXML & "<ow:body/>"
            Else
                ToXML = ToXML & "<ow:body>" & Wikify(vText) & "</ow:body>"
            End If
        End If
        If pIncludeText = 2 Or pIncludeText = 3 Then
            ToXML = ToXML & "<ow:raw>" & PCDATAEncode(vText) & "</ow:raw>"
        End If

        If cAllowAttachments Then
            If IsObject(vAttachmentCollection) Then
                vCount = vAttachmentCollection.Count - 1
                If vCount >= 0 Then
                    ToXML = ToXML & "<ow:attachments>"
                    For i = 0 To vCount
                        ToXML = ToXML & vAttachmentCollection.ElementAt(i).ToXML(vName, "")
                    Next
                    ToXML = ToXML & "</ow:attachments>"
                End If
            End If
        End If

        ToXML = ToXML & "</ow:page>"
    End Function
End Class



Class Change
    Private vStatus, vRevision, vTimestamp, vMinorEdit, vBy, vByAlias, vComment
    Private vAttachmentChanges

    Private Sub Class_Initialize()
        vStatus = "new"
    End Sub

    Private Sub Class_Terminate()
        Set vAttachmentChanges = Nothing
    End Sub

    Public Property Get Status()
        Status = vStatus
    End Property

    Public Property Let Status(pStatus)
        Select Case pStatus
        Case 1
            vStatus = "new"
        Case 2
            vStatus = "updated"
        Case 3
            vStatus = "deleted"
        Case Else
            ' must never happen
            Response.Write("DING DONG !!!")
            vStatus = "unknown"
        End Select
    End Property

    Public Property Get Revision()
        Revision = vRevision
    End Property

    Public Property Let Revision(pRevision)
        vRevision = pRevision
    End Property

    Public Property Get Timestamp
        Timestamp = vTimestamp
    End Property

    Public Property Let Timestamp(pTimestamp)
        vTimestamp = pTimestamp
    End Property

    Public Property Get MinorEdit()
        MinorEdit = vMinorEdit
    End Property

    Public Property Let MinorEdit(pMinorEdit)
        If pMinorEdit = 1 Or pMinorEdit = "1" Or pMinorEdit = "true" Or pMinorEdit = "on" Then
            vMinorEdit = 1
        Else
            vMinorEdit = 0
        End If
    End Property

    Public Property Get By()
        By = vBy
    End Property

    Public Property Let By(pBy)
        If cMaskIPAddress Then
            vBy = s(pBy, "\.\d+$", ".xxx", False, True)
        Else
            vBy = pBy
        End If
    End Property

    Public Property Get ByAlias()
        ByAlias = vByAlias
    End Property

    Public Property Let ByAlias(pByAlias)
        vByAlias = pByAlias
    End Property

    Public Sub UpdateBy()
        If (GetRemoteUser() <> vBy) Then
            vStatus = "updated"
        End If
    End Sub

    Public Property Get Comment()
        Comment = vComment
    End Property

    Public Property Let Comment(pComment)
        vComment = pComment
    End Property

    Public Sub AddAttachmentChange(pAttachmentChange)
        If Not IsObject(vAttachmentChanges) Then
            Set vAttachmentChanges = New Vector
        End If
        vAttachmentChanges.Push(pAttachmentChange)
    End Sub

    Public Function ToXML()
        ToXML = ToXML & "<ow:change revision='" & vRevision & "' status='" & vStatus & "'"
        If vMinorEdit = 1 Then
            ToXML = ToXML & " minor='true'>"
        Else
            ToXML = ToXML & " minor='false'>"
        End If
        ToXML = ToXML & "<ow:by name='" & CDATAEncode(vBy) & "'"
        If vByAlias <> "" Then
            ToXML = ToXML & " alias='" & CDATAEncode(vByAlias) & "'>" & PCDATAEncode(PrettyWikiLink(vByAlias)) & "</ow:by>"
        Else
            ToXML = ToXML & "/>"
        End If
        ToXML = ToXML & "<ow:date>" & FormatDateISO8601(vTimestamp) & "</ow:date>"
        If vComment <> "" Then
            ToXML = ToXML & "<ow:comment>" & PCDATAEncode(vComment) & "</ow:comment>"
        End If
        If IsObject(vAttachmentChanges) Then
            Dim i
            For i = 0 To vAttachmentChanges.Count - 1
                ToXML = ToXML & vAttachmentChanges.ElementAt(i).ToXML()
            Next
        End If
        ToXML = ToXML & "</ow:change>"
    End Function
End Class


Class Attachment
    Private vName, vRevision, vHidden, vDeprecated, vFilename, vTimestamp, vFilesize, vBy, vByAlias, vComment

    Private Sub Class_Initialize()
    End Sub

    Private Sub Class_Terminate()
    End Sub

    Public Property Get Name()
        Name = vName
    End Property

    Public Property Let Name(pName)
        vName = pName
    End Property

    Public Property Get Revision()
        Revision = vRevision
    End Property

    Public Property Let Revision(pRevision)
        vRevision = pRevision
    End Property

    Public Property Get Hidden()
        Hidden = vHidden
    End Property

    Public Property Let Hidden(pHidden)
        vHidden = pHidden
    End Property

    Public Property Get Deprecated()
        Deprecated = vDeprecated
    End Property

    Public Property Let Deprecated(pDeprecated)
        vDeprecated = pDeprecated
    End Property

    Public Property Get Filename()
        Filename = vFilename
    End Property

    Public Property Let Filename(pFilename)
        vFilename = pFilename
    End Property

    Public Property Get Timestamp()
        Timestamp = vTimestamp
    End Property

    Public Property Let Timestamp(pTimestamp)
        vTimestamp = pTimestamp
    End Property

    Public Property Get Filesize()
        Filesize = vFilesize
    End Property

    Public Property Let Filesize(pFilesize)
        vFilesize = pFilesize
    End Property

    Public Property Get By()
        By = vBy
    End Property

    Public Property Let By(pBy)
        If cMaskIPAddress Then
            vBy = s(pBy, "\.\d+$", ".xxx", False, True)
        Else
            vBy = pBy
        End If
    End Property

    Public Property Get ByAlias()
        ByAlias = vByAlias
    End Property

    Public Property Let ByAlias(pByAlias)
        vByAlias = pByAlias
    End Property

    Public Property Get Comment()
        Comment = vComment
    End Property

    Public Property Let Comment(pComment)
        vComment = pComment
    End Property

    Private Function GetIcon()
        Dim vPos, vExtension
        vPos = InStrRev(vName, ".")
        If vPos > 0 Then
            vExtension = Mid(vName, vPos + 1)
            If Not m(vExtension, "(" & gDocExtensions & ")", True, True) Then
                vExtension = "empty"
            End If
        Else
            vExtension = "empty"
        End If
        GetIcon = vExtension
    End Function

    Private Function FormatSize(pSize)
        FormatSize = Int((pSize / 1000) + 1)
        If FormatSize >= 1000000 Then
            FormatSize = Int((FormatSize / 1000000) + 1) & "," & (Int((FormatSize / 1000) + 1) - Int((FormatSize / 1000000) + 1)) & "," & (FormatSize Mod 1000)
        Elseif FormatSize >= 1000 Then
            FormatSize = Int((FormatSize / 1000) + 1) & "," & (FormatSize Mod 1000)
        End If
    End Function

    Public Function ToLinkXML(pHref, pText)
        ToLinkXML = "<ow:link name='" & CDATAEncode(vName) & "'" _
                  & " href='" & pHref & "' date='" & FormatDateISO8601(vTimestamp) & "'" _
                  & " attachment='true'>" _
                  & PCDATAEncode(pText) & "</ow:link>"
    End Function

    Public Function ToXML(pPagename, pText)
        Dim vAttachmentLink, vIsImage
        ToXML = ToXML & "<ow:attachment name='" & CDATAEncode(vName) & "' revision='" & vRevision & "' hidden='"
        If vHidden = 1 Then
            ToXML = ToXML & "true"
        Else
            ToXML = ToXML & "false"
        End If
        ToXML = ToXML & "' deprecated='"
        If vDeprecated = 1 Then
            ToXML = ToXML & "true"
        Else
            ToXML = ToXML & "false"
        End If
        ToXML = ToXML & "'>"
        vAttachmentLink = GetAttachmentLink(pPagename, vFilename)
        If m(vAttachmentLink, "\.(" & gImageExtensions & ")$", True, True) Then
            vIsImage = "true"
        Else
            vIsImage = "false"
        End If
        ToXML = ToXML & "<ow:file icon='" & getIcon() & "' size='" & FormatSize(vFilesize) & "' href='" & CDATAEncode(vAttachmentLink) & "' image='" & vIsImage & "'>" & PCDATAEncode(vName) & "</ow:file>"
        ToXML = ToXML & "<ow:by name='" & CDATAEncode(vBy) & "'"
        If vByAlias <> "" Then
            ToXML = ToXML & " alias='" & CDATAEncode(vByAlias) & "'>" & PCDATAEncode(PrettyWikiLink(vByAlias)) & "</ow:by>"
        Else
            ToXML = ToXML & "/>"
        End If
        ToXML = ToXML & "<ow:date>" & FormatDateISO8601(vTimestamp) & "</ow:date>"
        If vComment <> "" Then
            ToXML = ToXML & "<ow:comment>" & PCDATAEncode(vComment) & "</ow:comment>"
        End If
        ToXML = ToXML & PCDATAEncode(pText)
        ToXML = ToXML & "</ow:attachment>"
    End Function
End Class


Class AttachmentChange
    Private vName, vRevision, vTimestamp, vBy, vByAlias, vAction

    Private Sub Class_Initialize()
    End Sub

    Private Sub Class_Terminate()
    End Sub

    Public Property Get Name()
        Name = vName
    End Property

    Public Property Let Name(pName)
        vName = pName
    End Property

    Public Property Get Revision()
        Revision = vRevision
    End Property

    Public Property Let Revision(pRevision)
        vRevision = pRevision
    End Property

    Public Property Get Timestamp()
        Timestamp = vTimestamp
    End Property

    Public Property Let Timestamp(pTimestamp)
        vTimestamp = pTimestamp
    End Property

    Public Property Get By()
        By = vBy
    End Property

    Public Property Let By(pBy)
        If cMaskIPAddress Then
            vBy = s(pBy, "\.\d+$", ".xxx", False, True)
        Else
            vBy = pBy
        End If
    End Property

    Public Property Get ByAlias()
        ByAlias = vByAlias
    End Property

    Public Property Let ByAlias(pByAlias)
        vByAlias = pByAlias
    End Property

    Public Property Get Action()
        Action = vAction
    End Property

    Public Property Let Action(pAction)
        vAction = pAction
    End Property

    Public Function ToXML()
        ToXML = ToXML & "<ow:attachmentchange name='" & CDATAEncode(vName) & "' revision='" & vRevision & "'>"
        ToXML = ToXML & "<ow:by name='" & CDATAEncode(vBy) & "'"
        If vByAlias <> "" Then
            ToXML = ToXML & " alias='" & CDATAEncode(vByAlias) & "'>" & PCDATAEncode(PrettyWikiLink(vByAlias)) & "</ow:by>"
        Else
            ToXML = ToXML & "/>"
        End If
        ToXML = ToXML & "<ow:date>" & FormatDateISO8601(vTimestamp) & "</ow:date>"
        If vAction <> "" Then
            ToXML = ToXML & "<ow:action>" & PCDATAEncode(vAction) & "</ow:action>"
        End If
        ToXML = ToXML & "</ow:attachmentchange>"
    End Function
End Class

%>