<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 and Session("MailSender") then
Dim Conn, rs, i, Today, CanChanceDate, HTMLCode
Dim DaySend, MonthSend, YearSend, HourSend, HourToSender
Dim IMG, DateSend, HourToSend, CountHoursToSend, HoursToSend
Dim TitleDate

DateSend = CheckTxt(request.form("DateSend")) 
DaySend = CheckNum(request.form("DaySend"))
MonthSend = CheckNum(request.form("MonthSend"))
YearSend = CheckNum(request.form("YearSend"))

TitleDate = DaySend
Select Case MonthSend
Case 1
		 TitleDate = TitleDate & " de Enero de "  
Case 2
		 TitleDate = TitleDate & " de Febrero de "
Case 3
		 TitleDate = TitleDate & " de Marzo de "		 
Case 4
		 TitleDate = TitleDate & " de Abril de "
Case 5
		 TitleDate = TitleDate & " de Mayo de "
Case 6
		 TitleDate = TitleDate & " de Junio de "
Case 7
		 TitleDate = TitleDate & " de Julio de "
Case 8
		 TitleDate = TitleDate & " de Agosto de "
Case 9
		 TitleDate = TitleDate & " de Septiembre de "
Case 10
		 TitleDate = TitleDate & " de Octubre de "
Case 11
		 TitleDate = TitleDate & " de Noviembre de "
Case 12
		 TitleDate = TitleDate & " de Diciembre de "
End Select
TitleDate = TitleDate & YearSend

if isDate(DateSend) and DaySend <> 0 and MonthSend <> 0 and YearSend <> 0 then
	 OpenConn Conn
	 'REsponse.write "select TimeToSend, Status, Subject from Messages where DateToSend=#" & YearSend & "/" & MonthSend & "/" & DaySend & "# order by TimeToSend Asc"
	 Set rs = Conn.Execute("select TimeToSend, Status, Subject from Messages where DateToSend=#" & YearSend & "/" & MonthSend & "/" & DaySend & "# order by TimeToSend Asc")
	 if Not rs.EOF then
	 		HoursToSend = rs.GetRows
	 CountHoursToSend = UBound(HoursToSend, 2)
	 end if
	 closeOBJ rs
	 CloseOBJ Conn
	 Today = Day(Now)

	 CanChanceDate = False
	 CountHoursToSend = -1
	 IMG = "<IMG src='../img/transparente.gif' width=160 height=12 border=0>"
	 'verificando que el dia aun puede agendar, esto sucede si es el dia es el actual o futuro
	 if DaySend > Today then
	 		CanChanceDate = True 
	 end if
	 
	 Response.write CanChanceDate
	 for i = 0 to 23
			 		select case i
					case 0
							 HourToSend = "12:00</b><SPAN class=labela>&nbsp;am</SPAN>"
							 HourToSender = "12:00 AM"
					case 1,2,3,4,5,6,7,8,9
							 HourToSend = "0" & i & ":00</b><SPAN class=labela>&nbsp;am</SPAN>"
							 HourToSender = "0" & i & ":00 AM"
					case 10,11
							 HourToSend = i & ":00</b><SPAN class=labela>&nbsp;am</SPAN>"
							 HourToSender = i & ":00 AM"
					case 12 
					  	 HourToSend = i & ":00</b><SPAN class=labela>&nbsp;pm</SPAN>"
							 HourToSender = i & ":00 PM"
					case 13, 14, 15, 16, 17, 18, 19, 20, 21  
					  	 HourToSend = "0" & i-12 & ":00</b><SPAN class=labela>&nbsp;pm</SPAN>"
							 HourToSender = "0" & i-12 & ":00 PM"
					case 22, 23  
					  	 HourToSend = i-12 & ":00</b><SPAN class=labela>&nbsp;pm</SPAN>"
							 HourToSender = i-12 & ":00 PM"
					end select
			 HTMLCode = HTMLCode & "<TR><TD width='23%' class=labela bgcolor=#FFCC99><b>" & HourToSend & _
			  					"<TD class=labela bgcolor=#FFFFFF>"
			 if CountHoursToSend >= 0 then
			 		if (i * 10000) = HoursToSend(i,0) Then
			 			 Select Case HoursToSend(i,1)
						 Case 0 'Mensaje agendado pero no enviado aun.
							 		IMG = "<IMG src='../img/transparente.gif' width=160 height=12 border=0>"
						 Case 1 'Mensaje que actualmente se esta agendando.
							 		IMG = "<IMG src='../img/transparente.gif' width=160 height=12 border=0>"
					   Case 2 'Mensaje que ya ha sido enviado.
							 	  IMG = "<IMG src='../img/transparente.gif' width=160 height=12 border=0>"
					   Case 3 'Mensaje enviado pero presento errores.
							    IMG = "<IMG src='../img/transparente.gif' width=160 height=12 border=0>"
					   End Select
							    IMG = IMG & HoursToSend(i,2)					
			    end if
			 end if 
			 if CanChanceDate then
			 		 HTMLCode = HTMLCode & "<A href='JavaScript:top.principal.agendarMensaje('" & DateSend & " " & HourToSender & "');'>" & IMG & "</A></TD></TR>" 
			 else
			 		 HTMLCode = HTMLCode & "</TD></TR>"
			 end if  
	 next
%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language=javascript src="img/config.js"></SCRIPT>
<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>
<SCRIPT language="javascript" src="img/validaciones.js"></SCRIPT>
<style type="text/css"><!--
  body, td  { font-family: arial; font-size: x-small; }
  a         { color: #0000BB; text-decoration: none; }
  a:hover   { color: #FF0000; text-decoration: underline; }
  .headline { font-family: arial black, arial; font-size: 28px; letter-spacing: -1px; }
  .headline2{ font-family: verdana, arial; font-size: 12px; }
  .subhead  { font-family: arial, arial; font-size: 18px; font-weight: bold; font-style: italic; }
  .backtotop     { font-family: arial, arial; font-size: xx-small;  }
  .code     { background-color: #EEEEEE; font-family: Courier New; font-size: x-small;
              margin: 5px 0px 5px 0px; padding: 5px;
              border: black 1px dotted;
            }
  font { font-family: arial black, arial; font-size: 28px; letter-spacing: -1px; }
--></style>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY bgcolor="#FF7900" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="JavaScript:self.focus()">
	<TABLE  border=0 cellspacing=1 cellpadding=1 width=100%>
		<TR><TD class=labela bgcolor=#FFCC99 align=center colspan=2><b><%=TitleDate%></b></TD></TR>
		<%=HTMLCode%>
	</TABLE>	
</BODY>
<a 
<%
end if
end if
%>
