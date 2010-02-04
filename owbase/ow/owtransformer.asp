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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owtransformer.asp,v $
'    $Revision: 1.6 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

Class Transformer
    Private vXmlDoc, vXslDoc, vXslTemplate, vXslProc, vIsIE

'    Public Property Let MSXML_VERSION(pMSXML_VERSION)
'        vMSXML_VERSION = pMSXML_VERSION
'    End Property
'
'    Public Property Let Cache(pCacheXSL)
'        vCacheXSL = pCacheXSL
'    End Property
'
'    Public Property Let StylesheetsDir(pDir)
'        vStylesheetsDir = pDir
'    End Property
'
'    Public Property Let Encoding(pEncoding)
'        vEncoding = pEncoding
'    End Property
'
'    Public Property Let WriteToOutput(pWriteToOutput)
'        vWriteToOutput = pWriteToOutput
'    End Property

    Private Sub Class_Initialize()
        On Error Resume Next
        If MSXML_VERSION = 4 Then
            Set vXmlDoc = Server.CreateObject("Msxml2.FreeThreadedDOMDocument.4.0")
        Else
            Set vXmlDoc = Server.CreateObject("Msxml2.FreeThreadedDOMDocument")
        End If
        vXmlDoc.async = False
        vXmlDoc.preserveWhiteSpace = True
        If Not IsObject(vXmlDoc) Then
            ' As this is the first time we try to instantiate the XML Doc object
            ' let's assume the user hasn't configured his/her owconfig file
            ' correctly yet. Switch MS XML Version to try again.
            If MSXML_VERSION = 4 Then
                MSXML_VERSION = 3
            Else
                MSXML_VERSION = 4
            End If
            If MSXML_VERSION = 4 Then
                Set vXmlDoc = Server.CreateObject("Msxml2.FreeThreadedDOMDocument.4.0")
            Else
                Set vXmlDoc = Server.CreateObject("Msxml2.FreeThreadedDOMDocument")
            End If
            vXmlDoc.async = False
            vXmlDoc.preserveWhiteSpace = True
            If Not IsObject(vXmlDoc) Then
                EndWithErrorMessage()
            Elseif MSXML_VERSION = 3 Then
                Response.Write("<b>WARNING:</b>You've configured your OpenWiki to use the MSXML v4 component, but you don't appear to have this installed. The application now falls back to use the MSXML v3 component. Please update your config file (usually file owconfig_default.asp) or install MSXML v4.<br />")
            Else
                Response.Write("<b>WARNING:</b>You've configured your OpenWiki to use the MSXML v3 component, but you don't appear to have this installed. The application now falls back to use the MSXML v4 component. Please update your config file (usually file owconfig_default.asp) or install MSXML v3.<br />")
            End If
        End If

        Dim vTemp
        vTemp = Request.ServerVariables("HTTP_USER_AGENT")
        If (InStr(vTemp, " MSIE 5.5;") > 0) Or (InStr(vTemp, " MSIE 6") > 0) Then
            vIsIE = True
        Else
            vIsIE = False
        End If
    End Sub

    Private Sub Class_Terminate()
        Set vXslProc = Nothing
        Set vXslTemplate = Nothing
        Set vXslDoc = Nothing
        Set vXmlDoc = Nothing
    End Sub

    Private Sub EndWithErrorMessage()
        Response.Write("<h2>Error: Missing MSXML Parser 3.0 Release</h2>")
        Response.Write("In order for this script to work correctly the component " _
                     & "MSXML Parser 3.0 Release " _
                     & "or a higher version needs to be installed on the server. " _
                     & "You can download this component from " _
                     & "<a href=""http://msdn.microsoft.com/xml"">http://msdn.microsoft.com/xml</a>.")
        Response.End
    End Sub

    Public Sub LoadXSL(pFilename)
        On Error Resume Next
        Set vXslTemplate = Nothing
        vXslTemplate = ""
        If cCacheXSL = 1 Then
            If IsObject(Application("ow__" & pFilename)) Then
                Set vXslTemplate = Application("ow__" & pFilename)
            End If
        End If
        If Not IsObject(vXslTemplate) Then
            If MSXML_VERSION = 4 Then
                Set vXslDoc = Server.CreateObject("Msxml2.FreeThreadedDOMDocument.4.0")
            Else
                Set vXslDoc = Server.CreateObject("Msxml2.FreeThreadedDOMDocument")
            End If
            vXslDoc.async = False
            If Not vXslDoc.load(Server.MapPath(OPENWIKI_STYLESHEETS & pFilename)) Then
                Response.Write("<p><b>Error in " & pFilename & ":</b> " & vXslDoc.parseError.reason & " line: " & vXslDoc.parseError.Line & " col: " & vXslDoc.parseError.linepos & "</p>")
                Response.End
            End If
            If MSXML_VERSION = 4 Then
                Set vXslTemplate = Server.CreateObject("Msxml2.XSLTemplate.4.0")
            Else
                Set vXslTemplate = Server.CreateObject("Msxml2.XSLTemplate")
            End If
            If Not IsObject(vXslTemplate) Then
                EndWithErrorMessage()
            End If
            Set vXslTemplate.stylesheet = vXslDoc
            If Err.Number <> 0 Then
                Response.Write("<p><b>Error in an included stylesheet</p>")
                Response.End
            End If
            If cCacheXSL Then
                Set Application("ow__" & pFilename) = vXslTemplate
            End If
        End If
        Set vXslProc = vXslTemplate.createProcessor()
        If Not IsObject(vXslProc) Then
            EndWithErrorMessage()
        End If
        On Error Goto 0
    End Sub

    Public Function TransformXmlDoc(pXmlDoc, pXslFilename)
        LoadXSL(pXslFilename)
        vXslProc.input = pXmlDoc
        vXslProc.transform
        TransformXmlDoc = vXslProc.output
    End Function

    Public Function Transform(pXmlStr)
        Transform = TransformXmlStr(pXmlStr, "ow.xsl")
    End Function

    Public Function TransformXmlStr(pXmlStr, pXslFilename)
        Dim vXmlStr
        vXmlStr = "<?xml version='1.0' encoding='" & OPENWIKI_ENCODING & "'?>" & vbCRLF _
                & gNamespace.ToXML(pXmlStr)

        'Response.ContentType = "text/html"
        'Response.Write(vXmlStr)
        'Response.Write(Server.HTMLEncode(vXmlStr) & "<br /><br />" & vbCRLF & vbCRLF)
        'Response.End

        If gAction = "xml" Or InStrRev(Request.QueryString, "&xml=1") > 0 Then
            If vIsIE Or gAction = "xml" Then
                Response.ContentType = "text/xml; charset:" & OPENWIKI_ENCODING & ";"
                Response.Write(vXmlStr)
                Response.End
            Else
                pXslFilename = "xmldisplay.xsl"
            End If
        End If

        If Not vXmlDoc.loadXML(vXmlStr) Then
            Response.ContentType = "text/html; charset:" & OPENWIKI_ENCODING & ";"
            Response.Write("<html><body><b>Invalid XML document</b>:<br /><br />")
            Response.Write(vXmlDoc.parseError.reason & " line: " & vXmlDoc.parseError.Line & " col: " & vXmlDoc.parseError.linepos)
            Response.Write("<br /><br /><hr />")
            Response.Write("<pre>" & Replace(Server.HTMLEncode(vXmlStr), vbCRLF, "<br />") & "</pre>")
            Response.Write("</body></html>")
        Else
            LoadXSL(pXslFilename)
            vXslProc.input = vXmlDoc
            vXslProc.transform

            TransformXmlStr = vXslProc.output

            If cEmbeddedMode = 0 Then
                If gAction = "edit" Then
                    Response.ContentType = "text/html; charset:" & OPENWIKI_ENCODING & ";"
                    Response.Expires = 0   ' expires in a minute
                Elseif gAction = "rss" Then
                    Response.ContentType = "text/xml; charset:" & OPENWIKI_ENCODING & ";"
                Else
                    Response.ContentType = "text/html; charset:" & OPENWIKI_ENCODING & ";"
                    Response.Expires = -1  ' expires now
                    'Response.ExpiresAbsolute = Now() - 1
                    'Response.AddHeader "Cache-Control", "must-revalidate"
                    Response.AddHeader "Cache-Control", "no-cache"
                End If
                Response.Write(TransformXmlStr)
            End If

        End If
    End Function

End Class
%>