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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owregexp.asp,v $
'    $Revision: 1.3 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'
' These functions simulate the m and s operations as available in the
' programming language perl. You can usually literally copy perl regular
' expressions and expect them to work with these functions.
'
' In perl you can do something like:
'
'     s/A(.*?)B(.*?)C/&MyMethod($1, $2)/ge
'
' When the match is made this will call the sub "MyMethod", pass it
' the two matched variables, and finally the match is substituted
' by whatever the sub returns.
'
' The function s below can behave in a similar manner. The perl expression
' shown above would be written as:
'
'     myText = s(myText, "A(.*?)B(.*?)C", "&MyMethod($1, $2)", True, True)
'
' In ASP the sub must return the value to be substituted via the global
' parameter sReturn. E.g.
'
'     Sub MyMethod(pParam1, pParam2)
'         If pParam1 = pParam2 Then
'             sReturn = "Same"
'         Else
'             sReturn = "Different"
'         End If
'     End Sub
'

' Global register which subs (called by s) should use to return their value.
Dim sReturn

' Reuse regular expression object
On Error Resume Next
Dim gRegEx
Set gRegEx = New RegExp
If Not IsObject(gRegEx) Then
    Response.Write("<h2>Error:</h2><p>Probable cause: Registry permission problem.</p>")
    Response.Write("This is a known problem with Microsoft.<br />" _
                 & "You can find more information about this problem in the following  " _
                 & "<a href=""http://support.microsoft.com/support/kb/articles/Q274/0/38.ASP"">Microsoft knowledge base article</a>.")
    Response.End
End If
On Error Goto 0


Function m(pText, pPattern, pIgnoreCase, pGlobal)
    If IsNull(pText) Then
        m = False
        Exit Function
    End If
    gRegEx.IgnoreCase = pIgnoreCase
    gRegEx.Global     = pGlobal
    gRegEx.Pattern    = pPattern
    m = gRegEx.Test(pText)
End Function


Function s(pText, pSearchPattern, pReplacePattern, pIgnoreCase, pGlobal)
    'Response.Write("<br /><br />Text: " & Server.HTMLEncode(pText))
    'Response.Write("<br />Patterns: " & Server.HTMLEncode(pSearchPattern) & " --> " & Server.HTMLEncode(pReplacePattern))

    If IsNull(pText) Then
        s = ""
        Exit Function
    End If

    gRegEx.IgnoreCase = pIgnoreCase
    gRegEx.Global     = pGlobal
    gRegEx.Pattern    = pSearchPattern
    If (Left(pReplacePattern, 1) <> "&") Then
        s = gRegEx.Replace(pText, pReplacePattern)
    Else
        Dim vText, vPrevLastIndex, vPrevNewPos
        Dim vMatch, vMatches, vSubMatch, i, vCmd, vReplacement

        vText          = pText
        vPrevLastIndex = 0
        vPrevNewPos    = 0

        pReplacePattern = Mid(pReplacePattern, 2)

        Set vMatches = gRegEx.Execute(pText)
        For Each vMatch In vMatches
            vCmd = pReplacePattern

            i = 0
            For Each vSubMatch in vMatch.SubMatches
                'Response.Write("<br />SubMatch: " & Server.HTMLEncode(vSubMatch))
                vCmd = Replace(vCmd, "$" & (i + 1), """" & Replace(vSubMatch, """", """""") & """")
                i = i+ 1
            Next

            'Response.Write("<br />REGEXP CMD: " & Server.HTMLEncode(vCmd))

            sReturn = ""
            vCmd = Replace(vCmd, vbCRLF, """ & vbCRLF & """)
            Execute("Call " & vCmd)
            vReplacement = sReturn

            ' replace vMatch.Value in vText by vReplacement
            vPrevNewPos = vPrevNewPos + (vMatch.FirstIndex - vPrevLastIndex)
            vText = Mid(vText, 1, vPrevNewPos) & vReplacement & Mid(vText, vPrevNewPos + vMatch.Length + 1)
            vPrevNewPos = vPrevNewPos + Len(vReplacement) + 1
            vPrevLastIndex = vMatch.FirstIndex + vMatch.Length + 1
        Next
        s = vText
    End If
End Function

%>
