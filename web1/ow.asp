<%@ Language=VBScript EnableSessionState=False %>
<%
Option Explicit
%>

<!-- #include virtual="ow/owall.asp" //-->

<%
' assume /ow/ is a virtual directory pointing to <install-dir>/owbase/ow/

' override only the default configuration settings.
' the defaults are in file ../owbase/ow/owconfig_default.asp

OPENWIKI_DB = "Driver={Microsoft ODBC for Oracle};Server=MYORAINST;Uid=ow;Pwd=owpwd;"
OPENWIKI_DB_SYNTAX = DB_ORACLE

OPENWIKI_TITLE           = "My Web"           ' title of your wiki
OPENWIKI_TIMEZONE        = "-09:00"           ' Timezone of the server running this wiki, valid values are e.g. "+04:00", "-09:00", etc.

%>

<%
OwProcessRequest()
%>