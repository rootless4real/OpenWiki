' ---------------------------------------------------------------------------
' Creates or deletes the virtual directory http://localhost/OpenWiki
'
' Valid command line switches: -i -u -q
'   -i  Install virtual directory
'   -u  Remove virtual directory
'   -q  Run in quiet mode
'
' ---------------------------------------------------------------------------
' See http://www.microsoft.com/windows2000/en/advanced/iis/default.asp?url=/windows2000/en/advanced/iis/htm/asp/aore8v5e.htm

Option Explicit

Dim vPath,vName, vWshShell, vObjArgs, vInstall, vRemove, vQuiet, i

vName = "OpenWiki"

' get current path to folder
vPath = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\", 15) - 1)
vPath = vPath & "\owbase"

'Set WshShell = WScript.CreateObject("WScript.Shell")
'vPath = WshShell.RegRead("HKLM\Software\OpenWiki") & "\owbase"

vQuiet = False
vInstall = False
vRemove = False

Set vObjArgs = WScript.Arguments

For i = 0 To vObjArgs.Count - 1
    If InStr(LCase(vObjArgs(i)), "i") <> 0 Then
        vInstall = True
    Elseif InStr(LCase(vObjArgs(i)), "u") <> 0 Then
        vRemove = True
    End If
    If InStr(LCase(vObjArgs(i)), "q") <> 0 Then
        vQuiet = True
    End If
Next

If vRemove Then
    DeleteVDir vName
Elseif vInstall Then
    CreateVDir vName, vPath
End If


Sub CreateVDir(pName, pPath)
    Dim vRoot, vBaseDir, vWebSite, ipSecurityObj, ipList

    On Error Resume Next

    ' get the local host default web
    Set vWebSite = GetObject("IIS://localhost/w3svc/1")
    If Not IsObject(vWebSite) Then
        Display "Unable to locate the Default Web Site. IIS must be installed and running."
        Exit Sub
    Else
        'Display vWebSite.name
    End if

    ' get the root
    Set vRoot = vWebSite.GetObject("IIsWebVirtualDir", "Root")
    If (Err <> 0) Then
        Display "Unable to access root for " & vWebSite.ADsPath
        Exit Sub
    Else
        'Display vRoot.name
    End If

    ' find or create the ow vroot.
    ' The physical path to http://localhost/ow is set to the ...<install-dir>\owbase folder
    Err.Number = 0 'Clear Error
    Set vBaseDir = GetObject(vRoot.ADsPath & "/" + pName)
    If Err.Number <> 0 Then
        Err.Number = 0
        Set vBaseDir = vRoot.Create("IIsWebVirtualDir", pName)
        vBaseDir.AccessRead = True
        vBaseDir.AccessFlags = 513  ' = 0x200 + 0x01 = MD_ACCESS_SCRIPT + MD_ACCESS_READ
        vBaseDir.AppCreate False
        vBaseDir.AspAllowSessionState = False
        vBaseDir.SetInfo

        ' This section restricts access to everyone except localhost (127.0.0.1).
        'Set ipSecurityObj = vBaseDir.IpSecurity
        'ipSecurityObj.GrantByDefault = False
        'ipList = ipSecurityObj.IPGrant
        'ReDim ipList(UBound(ipList) + 1)
        'ipList(UBound(ipList)) = "127.0.0.1"
        'ipSecurityObj.IPGrant = ipList
        'vBaseDir.IpSecurity = ipSecurityObj
        'vBaseDir.SetInfo

        vBaseDir.Path = pPath
        vBaseDir.AppFriendlyName = "OpenWiki"
        vBaseDir.SetInfo

        If (Err <> 0) Then
            Display "Unable to create " & vRoot.ADsPath & "/" & pName
            Exit Sub
        Else
            Err = 0
            'Display vBaseDir.name
        End If
    End If

End Sub


Sub DeleteVDir(pName)
    Dim vRoot, vBaseDir, vWebSite, ipSecurityObj, ipList
    On Error Resume Next

    ' get the local host default web
    Set vWebSite = GetObject("IIS://localhost/w3svc/1")
    If Not IsObject(vWebSite) Then
        Display "Unable to locate the Default Web Site. IIS must be installed and running."
        Exit Sub
    Else
        'Display vWebSite.name
    End If

    ' get the root
    Set vRoot = vWebSite.GetObject("IIsWebVirtualDir", "Root")
    If (Err <> 0) Then
        Display "Unable to access root for " & vWebSite.ADsPath
        Exit sub
    Else
        'display vRoot.name
    End If

    Err.Number = 0 'Clear Error
    Set vBaseDir = GetObject(vRoot.ADsPath)

    vBaseDir.Delete "IIsWebVirtualDir", pName
    vBaseDir.SetInfo

    If Not vQuiet Then
        WScript.Echo "Virtual directory http://localhost/" & vBaseDir.Name & "/" & pName & " deleted successfully."
    End If
End Sub


Sub Display(pMsg)
    If Not vQuiet Then
        WScript.Echo Now & ". Error Code: " & Hex(Err) & " - " & pMsg
    End If
End Sub

