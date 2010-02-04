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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owindex.asp,v $
'    $Revision: 1.3 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

Class IndexSchemes
    Private Sub Class_Initialize()
    End Sub

    Private Sub Class_Terminate()
    End Sub

    Public Function GetRecentChanges(pDays, pMaxNrOfChanges, pFilter, pShortVersion)
        Dim vList, vCount, i, j, vResult, vElem, vChange, vTimestamp
        If pMaxNrOfChanges > 0 Then
            vTimestamp = Now() - pDays
            Set vList = gNamespace.TitleSearch(".*", pDays, pFilter, 1, 1)
            vCount = vList.Count - 1
            For i = 0 To vCount
                Set vElem = vList.ElementAt(i)
                Set vChange = vElem.GetLastChange()
                If vChange.Timestamp > vTimestamp Then
                    vResult = vResult & vElem.ToXML(False)
                    j = j + 1
                    If j >= pMaxNrOfChanges Then
                        Exit For
                    End If
                End If
            Next
        End If
        GetRecentChanges = "<ow:recentchanges"
        If pFilter = 0 Or pFilter = 1 Then
            GetRecentChanges = GetRecentChanges & " majoredits='true'"
        Else
            GetRecentChanges = GetRecentChanges & " majoredits='false'"
        End If
        If pFilter = 0 Or pFilter = 2 Then
            GetRecentChanges = GetRecentChanges & " minoredits='true'"
        Else
            GetRecentChanges = GetRecentChanges & " minoredits='false'"
        End If
        If pShortVersion Then
            GetRecentChanges = GetRecentChanges & " short='true'"
        Else
            GetRecentChanges = GetRecentChanges & " short='false'"
        End If
        GetRecentChanges = GetRecentChanges & ">" & vResult & "</ow:recentchanges>"
    End Function

    Public Function GetTitleSearch(pPattern)
        Dim vList, i, vCount, vResult
        Set vList = gNamespace.TitleSearch(pPattern, 0, 0, 0, 0)
        vCount = vList.Count - 1
        For i = 0 To vCount
            vResult = vResult & vList.ElementAt(i).ToXML(False)
        Next
        GetTitleSearch = "<ow:titlesearch value='" & CDATAEncode(pPattern) & "' pagecount='" & gNamespace.GetPageCount() & "'>" & vResult & "</ow:titlesearch>"
    End Function

    Public Function GetFullSearch(pPattern, pIncludeTitles)
        Dim vList, i, vCount, vResult
        Set vList = gNamespace.FullSearch(pPattern, pIncludeTitles)
        vCount = vList.Count - 1
        For i = 0 To vCount
            vResult = vResult & vList.ElementAt(i).ToXML(False)
        Next
        GetFullSearch = "<ow:fullsearch value='" & CDATAEncode(pPattern) & "' pagecount='" & gNamespace.GetPageCount() & "'>" & vResult & "</ow:fullsearch>"
    End Function

    Public Function GetRandomPage(pNrOfPages)
        Dim vList, i, vCount, vIndex, vResult
        Set vList = gNamespace.TitleSearch(".*", 0, 0, 0, 0)
        Randomize
        vCount = vList.Count - 1
        For i = 1 To pNrOfPages
            vIndex = Int(vCount * Rnd)
            vResult = vResult & vList.ElementAt(vIndex).ToXML(False)
        Next
        GetRandomPage = "<ow:randompages>" & vResult & "</ow:randompages>"
    End Function

    Public Function GetTemplates(pPattern)
        Dim vList, i, vCount, vResult
        Set vList = gNamespace.TitleSearch(pPattern, 0, 0, 0, 0)
        vCount = vList.Count - 1
        For i = 0 To vCount
            vResult = vResult & vList.ElementAt(i).ToXML(False)
        Next
        GetTemplates = "<ow:templates>" & vResult & "</ow:templates>"
    End Function

    Public Function GetTitleIndex()
        Dim vList, vCount, i, vResult
        Set vList = gNamespace.TitleSearch(".*", 0, 0, 0, 0)
        vCount = vList.Count - 1
        For i = 0 To vCount
            vResult = vResult & vList.ElementAt(i).ToXML(False)
        Next
        GetTitleIndex = "<ow:titleindex>" & vResult & "</ow:titleindex>"
    End Function

    ' This function is pure crap! really really bad!
    ' needs a totally different implementation
    ' either needs an NT service or something similar that runs daily to
    ' generate the meta-data, or keep track of this meta-data when saving
    ' a new page.
    ' Also generate meta-data about concepts like TwinPages, MetaWiki, etc.
    Public Function GetWordIndex()
        Dim vList, vCount, i, j, vElem, vTitle, vWords, vValues, vRegEx, vMatches, vMatch, vKeys, vResult
        Dim vLast, vLastIndex
        Set vWords  = New Vector
        Set vValues = New Vector
        Set vRegEx  = New RegExp
        vRegEx.IgnoreCase = False
        vRegEx.Global = True
        vRegEx.Pattern = "[A-Z\xc0-\xde]+[a-z\xdf-\xff]+"
        Set vList = gNamespace.TitleSearch(".*", 0, 0, 0, 0)
        vCount = vList.Count
        For i = 0 To vCount - 1
            Set vElem = vList.ElementAt(i)
            vTitle = PrettyWikiLink(vElem.Name)
            Set vMatches = vRegEx.Execute(vTitle)
            For Each vMatch In vMatches
                vWords.Push(vMatch.Value)
                vValues.Push("<ow:word value='" & CDATAEncode(vMatch.Value) & "'>" & vElem.ToXML(False) & "</ow:word>")
            Next
        Next

        vCount = vWords.Count - 1
        For i = 0 To vCount
            vLast = "\xff\xff\xff\xff\xff"
            vLastIndex = 0
            For j = 0 To vCount
                If vWords.ElementAt(j) < vLast Then
                    vLast = vWords.ElementAt(j)
                    vLastIndex = j
                End If
            Next
            vWords.SetElementAt vLastIndex, "\xff\xff\xff\xff\xff"
            vResult = vResult & vValues.ElementAt(vLastIndex)
        Next

        Set vWords  = Nothing
        Set vValues = Nothing
        Set vRegEx  = Nothing
        GetWordIndex = "<ow:wordindex>" & vResult & "</ow:wordindex>"
    End Function

End Class
%>