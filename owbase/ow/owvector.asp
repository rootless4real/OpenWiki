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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owvector.asp,v $
'    $Revision: 1.2 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'
' Implements a resizable array.
'

Class Vector
    Private myStack
    Private myCount

    Private Sub Class_Initialize()
        Redim myStack(8)
        myCount = -1
    End Sub

    Private Sub Class_Terminate()
    End Sub

    Public Property Let Dimension(pDim)
        Redim myStack(pDim)
    End Property

    Public Property Get Count()
        Count = myCount + 1
    End Property

    Public Sub Push(pElem)
        myCount = myCount + 1
        If (UBound(myStack) < myCount) Then
            Redim Preserve myStack(UBound(myStack) * 2)
        End If
        Call SetElementAt(myCount, pElem)
    End Sub

    Public Function Pop()
        If IsObject(myStack(myCount)) Then
            Set Pop = myStack(myCount)
        Else
            Pop = myStack(myCount)
        End If
        myCount = myCount - 1
    End Function

    Public Function Top()
        If IsObject(myStack(myCount)) Then
            Set Top = myStack(myCount)
        Else
            Top = myStack(myCount)
        End If
    End Function

    Public Function ElementAt(pIndex)
        If IsObject(myStack(pIndex)) Then
            Set ElementAt = myStack(pIndex)
        Else
            ElementAt = myStack(pIndex)
        End If
    End Function

    Public Sub SetElementAt(pIndex, pValue)
        If IsObject(pValue) Then
            Set myStack(pIndex) = pValue
        Else
            myStack(pIndex) = pValue
        End If
    End Sub

    Public Sub RemoveElementAt(pIndex)
        Do While pIndex < myCount
            Call SetElementAt(pIndex, ElementAt(pIndex + 1))
            pIndex = pIndex + 1
        Loop
        myCount = myCount - 1
    End Sub

    Public Function IsEmpty()
        IsEmpty = (myCount < 0)
    End Function
End Class
%>