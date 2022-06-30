<%@ Language=VBScript %>
<%Option Explicit%>
<%
Dim MS
MS = Trim(Request.QueryString("MS"))
%>
<HTML>
<HEAD>
</HEAD>
<BODY onload="javascript:document.forma.submit();" BGCOLOR="FFFFFF">
<FORM name=forma method=post action=default.asp?MS=<%=MS%> target=_top>
</FORM>
</BODY>
</HTML>
