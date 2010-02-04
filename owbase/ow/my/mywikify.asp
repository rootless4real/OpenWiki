
<%

Sub MyInitLinkPatterns()

    ' add here any custom defined link patterns

    ' and/or change cq override the patterns defined by default in InitLinkPatterns.

End Sub


' Here you can define your own custom made Processing Instructions.
' See also http://openwiki.com/?HelpOnProcessingInstructions
Function MyWikifyProcessingInstructions(pText)

    ' example of dealing with a processing instruction
    If m(pText, "^#STOPWORDS\s+", False, False) Then
        ' Add every word following the #STOPWORDS PI to the gStopWords string
        ' All these words, when present in the current page, will NOT be hyperlinked.
        Dim vPos, vTemp
        vPos = InStr(pText, vbCR)
        If vPos > 0 Then
            vTemp = Trim(Mid(pText, 11, vPos - 11))
            If vTemp <> "" Then
                vTemp = s(vTemp, "\s+", "|", False, True)
                gStopWords = gStopWords & "|" & vTemp
            End If
            pText = Mid(pText, vPos + 1)
        End If
    End If

    ' process other processing instructions you'd like to create here

    MyWikifyProcessingInstructions = pText
End Function


Function MyMultiLineMarkupStart(pText)
    ' pText = s(pText, "<svg>([\s\S]*?)<\/svg>", "&StoreSVGML($1)", True, True)
    MyMultiLineMarkupStart = pText
End Function


Function MyMultiLineMarkupEnd(pText)
    ' The <comment> tag stores text that doesn't show up at all.
    ' Uncomment the next line if you want to support the <comment> tag
    ' pText = s(pText, "\&lt;comment\&gt;([\s\S]*?)\&lt;\/comment\&gt;", "", True, True)

    MyMultiLineMarkupEnd = pText
End Function


Function MyLastMinuteChanges(pText)

    MyLastMinuteChanges = pText
End Function

%>