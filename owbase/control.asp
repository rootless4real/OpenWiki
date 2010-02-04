<%@ Language=VBScript EnableSessionState=False %>
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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/control.asp,v $
'    $Revision: 1.3 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'
' This script shows an example of how to embed the OpenWiki software into an
' existing application in an easy way. Simply include the ow/owall.asp file
' and set the variable cEmbeddedMode to value 1. To transform the text, call
' the method TransformEmbedded. That's it!
'
' Note that in embedded mode the feature set is limited. Most macros won't
' work for example, nor will processing instructions. However, things like
' TableOfContents, Footnotes, automatic hyperlinking, headers, horizontal
' rules, code blocks, emoticons, lists, tables, etc. all work.
'
' Also note that WikiWords still work. If you want to disable this
' then set the variable cWikiLinks to value 0.
'
' If you get an error saying that a datasource could not be found, then
' set the variable OPENWIKI_DB to value "".
'
'
Option Explicit
%>

<!-- #include file="ow/owall.asp" //-->

<%
cEmbeddedMode = 1

' uncomment next line if you do not want to hyperlink WikiPages.
'cWikiLinks = 0

' uncomment next line if you do not have a database backing up the control.
'OPENWIKI_DB=""

'gServerRoot          = "http://www.mywikisite.com"
'OPENWIKI_SCRIPTNAME  = "/foo/ow/ow.asp"
%>

<html>
<head>
<link rel="stylesheet" type="text/css" href="ow/css/ow.css" />
</head>
<body>

<%
If Request("edit") = 1 Then
%>
    <form action="control.asp" method="post">
        <textarea name="txt1" cols="40" rows="12"><%=Request("txt1")%></textarea>
        <br />
        <textarea name="txt2" cols="40" rows="12"><%=Request("txt2")%></textarea>
        <br />
        <input type="submit" name="submit" value="Submit">
    </form>
<%
Else
%>
    <h1>Text 1</h1>
    <%= TransformEmbedded(Request("txt1")) %>

    <hr size="2" />

    <h1>Text 2</h1>
    <%= TransformEmbedded(Request("txt2")) %>

    <p />
    <form action="control.asp" method="post">
        <input type="hidden" name="txt1" value="<%=Request("txt1")%>">
        <input type="hidden" name="txt2" value="<%=Request("txt2")%>">
        <input type="hidden" name="edit" value="1">
        <input type="submit" name="submit" value="Edit">
    </form>
<%
End If
%>

</body>
</html>
