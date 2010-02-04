
<%

' Examples of custom build macros.
'
' When you create a new macro add the letter P to the name of
' the sub for each parameter you define. A macro can take at
' most 2 parameters.
'
' A macro should return the value that is supposed to be
' substituted in the text by setting the global variable
' gMacroReturn.

' For each macro you add below, you must add it's name to the
' return value of this function. Seperate the names by the
' pipe (|) character.
'
' If you want to redefine all available macros set the
' variable gMacros (see also owpatterns.asp).
Function MyMacroPatterns
    If cEmbeddedMode = 0 Then
        'gMacros = "BR|RecentChanges|TitleSearch|FullSearch|TextSearch|TableOfContents|WordIndex|TitleIndex|GoTo|RandomPage|InterWiki|SystemInfo|Include|PageCount|UserPreferences|Icon|Anchor|Date|Time|DateTime|Syndicate|Aggregate|Footnote"
        MyMacroPatterns = "Glossary|Files"
    End If
End Function


' code by Dan Rawsthorne
' taken from http://openwiki.com/?OpenWiki/Suggestions
Sub MacroGlossaryP(pParams)
    gMacroReturn = GetGlossaryP(pParams)
End Sub

Public Function GetGlossaryP(pPattern)
    Dim vList, vCount, i, vResult
    Set vList = gNamespace.FullSearch(pPattern, False)
    vCount = vList.Count - 1
    For i = 0 To vCount
        vResult = vResult & vList.ElementAt(i).ToXML(False)
    Next
    GetGlossaryP = "<ow:titleindex>" & vResult & "</ow:titleindex>"
End Function


' original code by Leopold Faschalek
' modified by Dave Cantrell
' modified by Lawrence Pit
' taken from http://openwiki.com/?OpenWiki/Suggestions
Sub MacroFilesP(pPath)
    Call MacroFilesPP(pPath, "[\s\S]*")
End Sub

Sub MacroFilesPP( pPath, pWild )
  dim oFso, oFolder, oFiles, oFile
  Set oFso = Server.CreateObject( "Scripting.FileSystemObject" )
  If oFso.FolderExists( pPath ) then
    Set oFolder = oFso.GetFolder( pPath )
    Set oFiles  = oFolder.Files
    'parses path and converts to UNC path so files can be retrieved from server across network
    dim sLocalPathPrefix  : sLocalPathPrefix = Left( oFolder.Path, 1 )
    dim sUncPath
    sUncPath = "file:\\\" & pPath & "\"
    'sUncPath = "\\mymachine\" & Lcase( sLocalPathPrefix ) & "$" & Right( oFolder.Path, Len( oFolder.Path ) - 2 ) & "\"  '"
    'sUncPath = "http:\\www.mysite.com\" & Lcase( sLocalPathPrefix ) & "$" & Right( oFolder.Path, Len( oFolder.Path ) - 2 ) & "\"  '"
    gMacroReturn = "<b>" & sUncPath & "</b><ul>"
    For each oFile in oFiles
      If m( oFile.Name, pWild, false, true ) then
        gMacroReturn = gMacroReturn & "<li><a href='" & sUncPath & oFile.Name & "' target='_blank'>" & oFile.Name & "</a></li>"
      End If
    Next
    gMacroReturn = gMacroReturn & "</ul>"
  Else
    gMacroReturn = "<ow:error>error in path: " & pPath & "</ow:error>"
  End If
  Set oFile   = Nothing
  Set oFiles  = Nothing
  Set oFolder = Nothing
  Set oFso    = Nothing

  ' prevent the caching of pages wherein this macro is used
  cCacheXML = False
End Sub



Sub MacroPagechangedP(pParam)
    Dim vPage, vTimestamp
    Set vPage = gNamespace.GetPage(pParam, 0, False, False)
    If vPage.Exists Then
        vTimestamp = vPage.GetLastChange().Timestamp()
        gMacroReturn = FormatDate(vTimestamp)
    End If
End Sub



%>