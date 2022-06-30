<%@ Language=VBScript%>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<!-- #INCLUDE file="MD5.asp" -->
<%
Dim Conn, rs, OperatorID, OperatorLogin, OperatorPwd, OperatorLevel, Menu, ProductCount, SystemDate
OperatorLogin = Request.Form("login")
OperatorPwd = Request.Form("pwd")
OperatorID = 0
	'Borrando Posibles Sesiones Antiguas
	If Trim(OperatorLogin) <> "" And Trim(OperatorPWD) <> "" Then
		OperatorPWD = MD5(OperatorPWD)
		openConn2 Conn
			Set rs = Conn.Execute("select id_usuario from usuarios_empresas where pw_name='" & OperatorLogin & "' and pw_passwd='" & OperatorPWD & "' order by id_usuario Asc")
			if Not rs.EOF Then
				OperatorID = Cint(rs(0))
			end if
		CloseOBJs rs, Conn
		 
		if OperatorID = 318 or OperatorID=332 then
			Session("Login") = OperatorLogin
			Session("OperatorID") = OperatorID
		End if
	End if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- saved from url=(0034)http://10.10.1.10/login/login.aspx -->
<HTML><HEAD><TITLE>eWMS - Consola Administrativa</TITLE>
<META content="text/html; charset=utf-8" http-equiv=Content-Type>
<META name=GENERATOR content="MSHTML 9.00.8112.16437">
<META name=CODE_LANGUAGE content="Visual Basic .NET 7.1">
<META name=vs_defaultClientScript content=JavaScript>
<META name=vs_targetSchema 
content=http://schemas.microsoft.com/intellisense/ie5>
<SCRIPT language=javascript src="imgwms/coolbuttons2.js"></SCRIPT>
<script>
<%if CheckNum(Session("OperatorID"))>0 then%>
	window.open('Lines.asp','Lines','height=150,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0,status=1');
<%end if%>
</script>
<META name=GENERATOR content="Microsoft Visual Studio .NET 7.1">
<META name=CODE_LANGUAGE content="Visual Basic .NET 7.1">
<META name=vs_defaultClientScript content=JavaScript>
<META name=vs_targetSchema 
content=http://schemas.microsoft.com/intellisense/ie5><LINK rel=stylesheet 
type=text/css 
href="imgwms/form.css"></HEAD>
<BODY onload=DoLogin() bottomMargin=50 
background=imgwms/background.jpg 
leftMargin=390 rightMargin=0 topMargin=250 MS_POSITIONING="GridLayout">
<DIV style="BEHAVIOR: url(../../../Users/miguel-urbina/Webservices/webservice.htc)" id=service></DIV>
<FORM id=forma method=post name=forma action=LogLines.asp><BR><BR>
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TBODY>
  <TR>
    <TD width=1></TD>
    <TD vAlign=bottom align=left>
      <TABLE border=0 cellSpacing=0 width="100%">
        <TBODY></TBODY></TABLE>
      <TABLE style="WIDTH: 549px; HEIGHT: 16px" border=0 cellSpacing=0 
      cellPadding=0 width=549>
        <TBODY>
        <TR>
          <TD style="WIDTH: 125px; HEIGHT: 1px" width=125>&nbsp;<FONT 
            style="FONT-FAMILY: verdana; COLOR: white; FONT-SIZE: 11pt"><STRONG>eWMS</STRONG> 
            </FONT></TD>
          <TD style="BORDER-LEFT: white thin solid; WIDTH: 43px; HEIGHT: 1px" 
          width=43>
            <P><FONT color=gainsboro>&nbsp;Usuario:</FONT></P></TD>
          <TD style="WIDTH: 265px; HEIGHT: 1px" vAlign=top width=265>
          <INPUT style="WIDTH: 41.25%; FONT-FAMILY: Verdana; HEIGHT: 19px; FONT-SIZE: 8pt" id=login name=login  size=15></TD></TR>
        <TR>
          <TD style="WIDTH: 125px; HEIGHT: 20px" width=125><FONT 
            size=1><STRONG>Modulo</STRONG> Administrativo</FONT></TD>
          <TD style="BORDER-LEFT: white thin solid; WIDTH: 43px; HEIGHT: 20px" 
          width=43>
            <P><FONT color=gainsboro>&nbsp;Password:</FONT></P></TD>
          <TD style="WIDTH: 265px; HEIGHT: 20px" vAlign=top width=265>
          <INPUT style="WIDTH: 41.25%; FONT-FAMILY: Verdana; HEIGHT: 19px; FONT-SIZE: 8pt" id=txtPassword name=pwd size=15 type=password> 
          &nbsp; <IMG 
            style="CURSOR: hand" onclick="Javascript:document.forma.submit();"
            src="imgwms/go.gif"></TD></TR>
        <TR>
          <TD colSpan=3 align=left>&nbsp; </TD></TR>
        <TR>
          <TD colSpan=3 align=left><FONT color=#ffffff 
            size=1>AIMAR&nbsp;|&nbsp;Enterprise Warehouse Management 
            System&nbsp;| Redes de Control, S.A.</FONT></TD></TR></TBODY></TABLE></TD>
    <TD width=1></TD></TR>
  <TR></TR></TBODY></TABLE><BR></FORM>
  </BODY>
  </HTML>
