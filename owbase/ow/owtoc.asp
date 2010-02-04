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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owtoc.asp,v $
'    $Revision: 1.2 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

Class TableOfContents
    Private vTOC, vTOCStart, vTOCCurLevel, myCount

    Private Sub Class_Initialize()
        vTOCStart = 0
        vTOCCurLevel = -1
        myCount = 0
    End Sub

    Private Sub Class_Terminate()
    End Sub

    Public Property Get Count()
        Count = myCount + 1
    End Property

    Public Sub AddTOC(pLevel, pStr)
        Dim i
        If vTOCStart = 0 Then
            vTOCStart = pLevel
            vTOCCurLevel = pLevel - 1
        End If
        Do While (vTOCCurLevel < pLevel)
            vTOC = vTOC & "<" & GetTOCElement(vTOCCurLevel) & ">" & vbCRLF
            vTOCCurLevel = vTOCCurLevel + 1
        Loop
        Do While (vTOCCurLevel > pLevel)
            vTOC = vTOC & "</" & GetTOCEndElement() & ">" & vbCRLF
            vTOCCurLevel = vTOCCurLevel - 1
        Loop
        Do While (vTOCStart > pLevel)
            vTOCStart = vTOCStart - 1
            vTOC = "<" & GetTOCElement(vTOCStart) & ">" & vTOC & vbCRLF
        Loop

        ' 25/02/2002 LP: commented because a <br/> after an <li> shows
        ' up bad in Opera. Besides, imo it's unnecessary.
        'If cNumTOC = 0 then
        '    vTOC = vTOC & pStr & "<br />" & vbCRLF
        'Else
        '    vTOC = vTOC & pStr & vbCRLF
        'End if
        vTOC = vTOC & pStr & vbCRLF

        myCount = myCount + 1

    End Sub


    Public Function GetTOC
        Do While (vTOCCurLevel >= vTOCStart)
            vTOC = vTOC & "</" & GetTOCEndElement() & ">" & vbCRLF
            vTOCCurLevel = vTOCCurLevel - 1
        Loop
        GetTOC = vTOC
    End Function

    Private Function GetTOCElement(pLevel)
        If cNumTOC = 0 then
            GetTOCElement = "dl"
        Elseif pLevel = 0 Then
            GetTOCElement = "ol"
        Elseif pLevel = 1 Then
            GetTOCElement = "ol type=""I"""
        Elseif pLevel = 2 Then
            GetTOCElement = "ol type=""a"""
        Elseif pLevel = 3 Then
            GetTOCElement = "ol type=""i"""
        Elseif pLevel = 4 Then
            GetTOCElement = "ol type=""1"""
        Else
            GetTOCElement = "ol"
        End If
    End Function

    Private Function GetTOCEndElement()
        If cNumTOC = 0 then
            GetTOCEndElement = "dl"
        Else
            GetTOCEndElement = "ol"
        End if
    End Function

End Class
%>