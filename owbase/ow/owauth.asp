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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owauth.asp,v $
'    $Revision: 1.2 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'


Function GetRemoteUser()
    Dim vPos
    GetRemoteUser = Request.ServerVariables("REMOTE_USER")
    If cStripNTDomain Then
        vPos = InStr(GetRemoteUser, "\")
        If vPos > 0 Then
            GetRemoteUser = Mid(GetRemoteUser, vPos + 1)
        End If
    End If
End Function

Function GetRemoteAlias()
    GetRemoteAlias = Request.Cookies(gCookieHash & "?up")("un")
End Function

' http://support.microsoft.com/support/kb/articles/Q245/5/74.ASP
'______________________________________________________________________________________________________________
Function GetRemoteHost()
    Dim vHost
    If cUseLookup Then
        vHost = Request.ServerVariables("REMOTE_HOST")
    End If
    If Not cUseLookup Or vHost = "" Then
        vHost = Request.ServerVariables("REMOTE_ADDR")
    End If
    GetRemoteHost = vHost
End Function

' you need administrator rights to do this
Sub EnableRemoteHostLookup(pCurrentWebOnly)
    Dim oIIS
    Dim vWebsite
    Dim vEnableRevDNS
    Dim vDisableRevDNS

    vEnableRevDNS = 1
    vDisableRevDNS = 0

    If pCurrentWebOnly Then
        Dim vPos
        vWebsite = Request.ServerVariables("INSTANCE_META_PATH")
        vPos = InStrRev(vWebsite, "/")
        If vPos > 0 Then
            vWebsite = "/" & Mid(vWebsite, vPos + 1) & "/ROOT"
        Else
            Exit Sub
        End If
    End If

    Set oIIS = GetObject("IIS://localhost/w3svc" & vWebsite)
    oIIS.Put "EnableReverseDNS", vEnableRevDNS
    oIIS.SetInfo
    Set oIIS = Nothing
End Sub

%>