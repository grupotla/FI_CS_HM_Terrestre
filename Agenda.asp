<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then

Dim DateSend, Today, HourToSend, ThisHour, FirstDate
Dim DaySend, MonthSend, YearSend, i, Subject, Action
Dim Label, DateOriginal, ObjectID

Action = Request("Action") 'opcion 1 = solo calendario, 2 = agenda de envio de mensajes
DateSend = Request("DateSend")
ObjectID = Request("OID")
Subject = Request("Subj") 'asunto del mensaje, en opcion 2
Label  = CheckTxt(Request("Label")) 'etiqueta del formulario en donde se colocara la fecha seleccionada
'REsponse.write DateSend & "<br>"
'DateSend = "14/01/2004 10:00 PM"
FirstDate = 0
if Not isDate(DateSend) then
		FirstDate = 1
		Today = Now+(Action-1)
		DaySend = Day(Today)
		MonthSend = Month(Today)
		YearSend = Year(Today)
		'12-01-2004 07:12 PM
		if DaySend < 10 then
			 DaySend = "0" & DaySend
		End if
		if MonthSend < 10 then
			 MonthSend = "0" & MonthSend		
		End if
		ThisHour = Hour(Time)
		select case ThisHour 
					case 0
							 HourToSend = "12:00 AM"
					case 1,2,3,4,5,6,7,8,9
							 HourToSend = "0" & ThisHour & ":00 AM"
					case 10,11
							 HourToSend = ThisHour & ":00 AM"
					case 12 
							 HourToSend = ThisHour & ":00 PM"
					case 13, 14, 15, 16, 17, 18, 19, 20, 21  
							 HourToSend = "0" & ThisHour-12 & ":00 PM"
					case 22, 23  
							 HourToSend = ThisHour-12 & ":00 PM"
		end select
		DateSend = Server.UrlEncode(DaySend & "/" & MonthSend & "/" & YearSend & " " & HourToSend)
else
		select case Action
		case 1
				 DateSend = Server.UrlEncode(DateSend & " 00:00 PM")
		case 2
				 DateSend = Server.UrlEncode(DateSend)
		end select 
end if
DateOriginal = DateSend
DateSend = DateSend & "&Subj=" & Server.URLEncode(Subject) & "&FD=" & FirstDate & "&Action=" & Action & "&Label=" & Label
%>
<HTML>
<HEAD>
	<TITLE>Agenda - Correo masivo - Administraci&oacute;n</TITLE>
	<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</HEAD>
<FRAMESET cols="*" rows="29,*" border=0>
	<FRAME src=Agenda_Dates.asp?DateSend=<%=DateSend%> name=principal scrolling=no>
	<FRAME src="blanco.html" name=calendario scrolling=no>
</FRAMESET><noframes></noframes>
</HTML>
<%end if%>
