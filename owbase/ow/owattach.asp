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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owattach.asp,v $
'    $Revision: 1.2 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'
'
'
'
'
'       !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'       !!! WARNING: YOU ARE POTENTIALLY RUNNING A SECURITY RISK !!!
'       !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
'
'
'
' Potentially people can upload files to your server that will be executed
' on your server when they point their browser to the uploaded file.
'
' This is possible because in IIS executables are mapped to file extensions.
' So for example, potentially someone can upload an .asp file, then point their
' browser to that file, after which the file is executed on your server, possibly
' wiping out your entire file system.
'
' In IIS you can view the application mappings by going to the properties of
' your website, choose tab "Home Directory", click on button "Configuration"
' and choose tab "App Mappings".
'
' Besides running a security risk on your server, by allowing people to upload
' files your users might unsuspectedly run into a mallicious file that is
' executed on their machines. For example, it's easy to upload an html file
' with very nasty javascript code; when your users view the html page as is,
' the javascript code will get executed on their machine, and your users might
' get very angry with you. ;)
'
' So in general: it's a bad idea to allow uploads on public websites. Only
' allow it when you trust your users.
'
' ---------------------------------------------------------------------------
'
' PLEASE READ THE BIG LETTERS OF THE OPENWIKI LICENSE AGAIN !
'
' It's printed above for your convenience. ;-)
'
' ---------------------------------------------------------------------------
'
' By default all files uploaded will get the extension .safe, except for files
' with extensions that are defined in the variable gDocExtensions (see file
' owpatterns.asp).
'
' If the variable gNotAcceptedExtensions is defined (see file owpattern.asp),
' then all files uploaded will keep their extensions, except the ones defined
' in the variable; they will still get the .safe extension. When you use this
' method you are advised to add all the extensions for which an application
' mapping is defined for the website in IIS.
'
' ---------------------------------------------------------------------------
'
' The is version 1 of the upload feature.
'
' This version of the upload feature assumes that you have installed the
' upload component ABC Upload version 4 from WebSupergoo. It's "free" and
' easy to install. See http://www.websupergoo.com/abcupload-1.htm.
'
' Using a different component would be quite easy if you know a bit of ASP.
' You'd only have to modify the sub ActionUpload below. Future versions of
' OpenWiki will probably support various upload components.
'
' This version saves the files to the file system. Saving to the database
' is again quite easy if you know a bit of ASP and databases, see comments
' inline the code below (note: you'd also need a separate script to retrieve
' the files). Future versions of OpenWiki will probably support storing
' files as BLOB's in the database.
'
' ---------------------------------------------------------------------------
'

Sub ActionAttach
    ActionView()
End Sub


Sub ActionUpload
    Response.Expires = -10000
    Server.ScriptTimeOut = OPENWIKI_UPLOADTIMEOUT
    Dim theForm, theField, vFilename
    On Error Resume Next
    Err.Number = 0
    Set theForm = Server.CreateObject("ABCUpload4.XForm")
    If Err.Number <> 0 Then
        Response.Write("<b>Error</b>: Missing component ABCUpload4. You can download this component from <a href='http://www.websupergoo.com/downloadftp.htm'>websupergoo.com</a>")
        Response.End
    End If
    On Error Goto 0
    theForm.MaxUploadSize = OPENWIKI_MAXUPLOADSIZE
    theForm.Overwrite = True
    theForm.AbsolutePath = False
    ' TODO: maybe implement pop-up progress-bar
    'theForm.ID = Request.QueryString("ID")

    'On Error Resume Next
    Set theField = theForm("file")(1)
    If theField.FileExists Then
        ' If you want to store your files as BLOBs in the database, then you should
        ' comment the next line
        CreateFolders()

        vFilename = theField.SafeFileName
        vFilename = gNamespace.SaveAttachmentMetaData(vFilename, theField.Length, theForm("link"), theForm("hide"), theForm("comment"))

        ' Save to filesystem.
        theField.Save OPENWIKI_UPLOADDIR & gPage & "/" & vFilename
    End If
    Response.Redirect(gScriptName & "?p=" & Server.URLEncode(gPage) & "&a=attach")
    Response.End
End Sub


Sub ActionHidefile
    Call gNamespace.HideAttachmentMetaData(Request("file"), Request("rev"), 1)
    Response.Redirect(gScriptName & "?p=" & Server.URLEncode(gPage) & "&a=attach")
    Response.End
End Sub

Sub ActionUndohidefile
    Call gNamespace.HideAttachmentMetaData(Request("file"), Request("rev"), 0)
    Response.Redirect(gScriptName & "?p=" & Server.URLEncode(gPage) & "&a=attach")
    Response.End
End Sub

Sub ActionTrashfile
    Call gNamespace.TrashAttachmentMetaData(Request("file"), Request("rev"), 1)
    Response.Redirect(gScriptName & "?p=" & Server.URLEncode(gPage) & "&a=attach")
    Response.End
End Sub

Sub ActionUndotrashfile
    Call gNamespace.TrashAttachmentMetaData(Request("file"), Request("rev"), 0)
    Response.Redirect(gScriptName & "?p=" & Server.URLEncode(gPage) & "&a=attach")
    Response.End
End Sub

Sub ActionAttachchanges
    Call gTransformer.Transform(gNamespace.GetPageAndAttachments(gPage, 0, False, True).ToXML(0))
    gActionReturn = True
End Sub



' If you want to store your files as BLOBs in the database, then you'd need
' to change this function.
'
' IN:
'   pPagename : page that has the attachment
'   pFilename : filename of the attachment
' RETURN: full URL to view/download the attachment
Function GetAttachmentLink(pPagename, pFilename)
    GetAttachmentLink = gServerRoot & OPENWIKI_UPLOADDIR & pPagename & "/" & pFilename
End Function


' Create all the subfolders if they do not exist yet.
Sub CreateFolders()
    Dim vFSO, vPosBegin, vPosEnd, vPath
    Set vFSO = Server.CreateObject("Scripting.FileSystemObject")
    If Not vFSO.FolderExists(Server.MapPath(OPENWIKI_UPLOADDIR & gPage & "/")) Then
        vPosBegin = 1
        vPath = Server.MapPath(OPENWIKI_UPLOADDIR)
        Do While True
            vPosEnd = InStr(vPosBegin, gPage, "/")
            If vPosEnd > vPosBegin Then
                vPath = vPath & "\" & Mid(gPage, vPosBegin, vPosEnd - vPosBegin)
                If Not vFSO.FolderExists(vPath) Then
                    Call vFSO.CreateFolder(vPath)
                End If
                vPosBegin = vPosEnd + 1
            Else
                vPath = vPath & "\" & Mid(gPage, vPosBegin)
                If Not vFSO.FolderExists(vPath) Then
                    Call vFSO.CreateFolder(vPath)
                End If
                Exit Do
            End If
        Loop
    End If
End Sub


%>