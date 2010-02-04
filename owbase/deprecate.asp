<%@ Language=VBScript EnableSessionState=False %>

<!-- #include file="ow/owpreamble.asp" //-->
<!-- #include file="ow/owconfig_default.asp" //-->
<!-- #include file="ow/owvector.asp" //-->
<!-- #include file="ow/owado.asp" //-->
<!-- #include file="ow/owattach.asp" //-->

<%

Dim gDaysToKeep

'
' This script deletes deprecated pages and attachments.
'

Response.Write("Look in the script! <p />")

' RUN AT YOUR OWN RISK !!!
' ALSO DELETES DEPRECATED ATTACHMENTS
' MAKE SURE YOU'VE SET THE VARIABLE OPENWIKI_DB CORRECTLY IN YOUR CONFIG FILE
'
' COMMENT THE NEXT LINE AND REFRESH THE PAGE
Response.End

' If a page is marked as deprecated, but was last modified less than
' <gDaysToKeep> days, then keep the page and/or attachment. Otherwise
' delete it.
gDaysToKeep = 30

Dim rs, q, v, vText
q = "SELECT wrv_name, wrv_timestamp, wrv_text FROM openwiki_revisions WHERE wrv_current = 1"
Set v = New Vector
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open q, OPENWIKI_DB, adOpenForwardOnly
Do While Not rs.EOF
    If rs("wrv_timestamp") < (Now() - gDaysToKeep) Then
        vText = rs("wrv_text")
        If Len(vText) >= 11 Then
            If Left(vText, 11) = "#DEPRECATED" Then
                Response.Write(rs("wrv_name") & "<br />")
                v.Push "" & rs("wrv_name")
            End If
        End If
    End If
    rs.MoveNext
Loop
Set rs = Nothing


Dim vFSO, vPagename
Set vFSO = Server.CreateObject("Scripting.FileSystemObject")

' delete pages and their attachments
Do While Not v.IsEmpty
    vPagename = v.Pop
    adoDMLQuery "DELETE FROM openwiki_revisions WHERE wrv_name = '" & Replace(vPagename, "'", "''") & "'"
    adoDMLQuery "DELETE FROM openwiki_attachments_log WHERE ath_wrv_name = '" & Replace(vPagename, "'", "''") & "'"
    adoDMLQuery "DELETE FROM openwiki_attachments WHERE att_wrv_name = '" & Replace(vPagename, "'", "''") & "'"
    If vFSO.FolderExists(Server.MapPath(OPENWIKI_UPLOADDIR & vPagename & "/")) Then
        vFSO.DeleteFolder(Server.MapPath(OPENWIKI_UPLOADDIR & vPagename & "/"))
    End If
Loop
Set v = Nothing

adoDMLQuery "DELETE FROM openwiki_pages WHERE NOT EXISTS (SELECT 'x' FROM openwiki_revisions WHERE wrv_name = wpg_name)"

' delete deprecated attachments
Dim vPath
q = "SELECT att_wrv_name, att_filename, att_timestamp FROM openwiki_attachments WHERE att_deprecated = 1"
Set v = New Vector
Set rs = Server.CreateObject("ADODB.Recordset")
'rs.Open q, OPENWIKI_DB, adOpenForwardOnly
rs.Open q, OPENWIKI_DB, adOpenKeyset, adLockOptimistic, adCmdText
Do While Not rs.EOF
    If rs("att_timestamp") < (Now() - gDaysToKeep) Then
        vPath = Server.MapPath(OPENWIKI_UPLOADDIR & rs("att_wrv_name") & "/" & rs("att_filename"))
        If vFSO.FileExists(vPath) Then
            Response.Write(vPath & "<br />")
            vFSO.DeleteFile(vPath)
            rs.Delete
        End If
    End If
    rs.MoveNext
Loop
Set rs = Nothing

Response.Write("<p />Done!")

Sub adoDMLQuery(pQuery)
    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open OPENWIKI_DB
    conn.Execute pQuery
    Set conn = Nothing
End Sub

%>