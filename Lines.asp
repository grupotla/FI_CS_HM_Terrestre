<%@ Language=VBScript %>
<!-- #INCLUDE file="utils.asp" -->
<%
	Dat = ConvertDate(now,3)
	Obs = Request.Form("Obs")
	Check Session("OperatorID")
	
	if Obs <> "" then
		OpenConn Conn	
		set rs = Conn.Execute("select CreatedDate, Obs from Agentsx where AgentID=10")
		if Not rs.EOF then
			OldDat = ConvertDate(rs(0),3)
			if Dat <> OldDat then
				OldObs = ""
			else
				OldObs = DC(rs(1),Dat)
			end if
		end if
		CloseOBJ rs
		
		'select case Session("OperatorID")
		'case 318
			OldObs = OldObs & "<font color=black>" & Ucase(Mid(Session("Login"),1,2)) & ">" & Obs & "</font><br>"
		'case else
		'	OldObs = OldObs & "<font color=gray>" & Ucase(Mid(Session("Login"),1,2)) & ">" & Obs & "</font><br>"
		'end select
		
		Conn.Execute("update Agentsx set Obs = '" & EC(OldObs,Dat) & "', CreatedDate='" & Dat & "' where AgentID=10")
		
		CloseOBJ Conn
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE></TITLE>
<META content="text/html; charset=utf-8" http-equiv=Content-Type>
<META name=GENERATOR content="MSHTML 9.00.8112.16437">
<META name=CODE_LANGUAGE content="Visual Basic .NET 7.1">
<META name=vs_defaultClientScript content=JavaScript>
<META name=vs_targetSchema 
content=http://schemas.microsoft.com/intellisense/ie5><LINK rel=stylesheet type=text/css href="imgwms/form.css">
<SCRIPT language=javascript src="imgwms/coolbuttons1.js"></SCRIPT>
</HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 scroll=no topMargin=0 bgColor=gainsboro>
<FORM name="forma" action="Lines.asp" method="post">
<TABLE border=0 cellSpacing=0 cellPadding=0 width="100%" height="100%">
  <TBODY>
  <TR>
    <TD height=30>
      <TABLE border=0 cellSpacing=0 cellPadding=0 width="100%">
        <TBODY>
        <TR>
          <TH bgColor=cornflowerblue width=10>&nbsp;</TH>
          <TH bgColor=cornflowerblue vAlign=middle width=40><IMG 
            src="imgwms/form.gif"></TH>
          <TH bgColor=cornflowerblue vAlign=middle align=left>&nbsp;</TH></TR></TBODY></TABLE></TD></TR><!--Toolbar-->
  <TR>
    <TD bgColor=black height=1></TD></TR>
  <TR>
    <TD bgColor=silver height=25>
      <TABLE border=0 cellSpacing=0 cellPadding=0 width="100%" height="100%">
        <TBODY>
        <TR>
          <TD>
            <TABLE cellSpacing=0 cellPadding=0 width="100%">
              <TBODY>
              <TR>
                <TD bgColor=silver width=1>&nbsp;</TD>
                <TD bgColor=darkblue width=1><IMG 
                  src="imgwms/clearpixel.gif"></TD>
                <TD bgColor=whitesmoke width=1><IMG 
                  src="imgwms/clearpixel.gif"></TD>
                <TD id=cmdNew class=coolButton title=New onClick="Javascript:document.forma.submit();" 
                bgColor=silver vAlign=middle width=40><IMG 
                  src="imgwms/new.gif"> </TD>
                <TD bgColor=darkblue width=1><IMG 
                  src="imgwms/clearpixel.gif"></TD>
                <TD bgColor=whitesmoke width=1><IMG 
                  src="imgwms/clearpixel.gif"></TD>
                <TD id=cmdDelete class=coolButton title=Delete onClick="Javascript:document.forma.Obs.value='';"; 
                bgColor=silver vAlign=middle width=40><IMG 
                  src="imgwms/delete.gif"> </TD>
                <TD bgColor=darkblue width=1><IMG 
                  src="imgwms/clearpixel.gif"></TD>
                <TD bgColor=whitesmoke width=1><IMG 
                  src="imgwms/clearpixel.gif"></TD>
                <TD 
      bgColor=silver>&nbsp;</TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD bgColor=black height=1></TD></TR><!--Data Fields-->
  <TR>
    <TD bgColor=gainsboro vAlign=top>
      <TABLE border=0 cellSpacing=0 cellPadding=0 width="100%">
        <TBODY>
        <TR>
          <TD height=3 width="5%">&nbsp;</TD>
          <TD width="25%"></TD>
          <TD width="65%"><IMG src="imgwms/clearpixel.gif"> </TD>
          <TD width="5%">&nbsp;</TD></TR>
        <TR>
          <TD width="5%">&nbsp;</TD>
          <TD width="25%">OBS_ESTADO:</TD>
          <TD width="65%">
          <iframe id="viewlines" style="background-color:gainsboro;" name="viewlines" src="http://10.10.1.21:8181/terrestre/admin/ViewLines.asp" frameborder="0" framespacing="0" scrolling="auto" width="249" height="40">
			Tu browser no soporta esta funcionalidad, favor contactar a soporte.
		   </iframe>
        </TD>
        </TR>
        <TR>
          <TD width="5%">&nbsp;</TD>
          <TD width="25%">OBS_ESTADO:</TD>
          <TD width="65%"><Textarea style="WIDTH: 85%; FONT-FAMILY: verdana; FONT-SIZE: 8pt" name="Obs" cols="30" rows="2"></Textarea></TD>
        </TR>
		    
            
            </TD>
      <TD width="5%">&nbsp;</TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
</FORM>
</BODY>
</HTML>
