<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Dim DateSend, Subject, FirstDate, Action, Label, DateOriginal, ObjectID
'DateOriginal = DateSend = (Request("DateOriginal"))
DateSend = (Request("DateSend"))
ObjectID = (Request("OID"))
Subject = (Request("Subj"))
Label = (Request("Label"))
FirstDate = (Request("FD"))
Action = (Request("Action"))
%>

<HTML>
<HEAD>
	<TITLE>Calendario - Administraci&oacute;n</TITLE>
	<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</HEAD>
<SCRIPT language="javascript">
//var dateOriginal = '<%=DateOriginal%>';
var dateField = '<%=DateSend%>';
var Subject = '<%=Subject%>';
var Action = <%=Action%>;
var Label = '<%=Label%>';
var ObjectID = '<%=ObjectID%>';
var FirstDate = <%=FirstDate%>;

document.calendaryFrame = top.calendario;
//document.clockFrame = top.hora;
document.numMonth = <%=Action%>;
document.textHelp = "Seleccione la hora y haga clic sobre el d&iacute;a";
document.patron = /(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2}) (AM|PM)/;
document.monthTemplate = "<TABLE border=0 cellspacing=0 cellpadding=1 width=128 height=100>\n<TR>";
document.monthTemplate += "	<TD width=128 class=calendary align=center colspan=7>{_MONTH_} {_YEAR_}</TD>\n";
document.monthTemplate += "{_DAYS_}\n</TABLE>\n";
document.weekTemplate = "  <TR>\n{_WEEK_}\n  </TR>\n";
document.headerTemplate = "	<TD width=20 bgcolor=#CCCCCC class=calendary align=center><b>{_DAY_}</b></TD>\n";
document.blankDayTemplate = "   <TD bgcolor=#FFFFFF class=calendary align=right>&nbsp;</TD>";
//document.dayTemplate = "	<TD bgcolor=#FFFFFF class=calendary align=right><A href=\"JavaScript:top.principal.updateDate('{_DAY_}-{_MONTH_}-{_YEAR_}')\" class=calendary>{_DAY_}</A></TD>";
document.dayTemplate = "	<TD bgcolor=#FFFFFF class=calendary align=right><A href=\"JavaScript:top.principal.agenda('{_DAY_}/{_MONTH_}/{_YEAR_}');top.principal.setDay('{_DAY_}/{_MONTH_}/{_YEAR_}')\" class=calendary>{_DAYNUM_}</A></TD>";
document.holiDayTemplate = document.dayTemplate;
document.selectDayTemplate = "	<TD bgcolor=#FF0000 class=calendary align=center>\n";
document.selectDayTemplate += "		<TABLE border=0 cellspacing=0 cellpadding=0 width=100%>\n";
document.selectDayTemplate += "		<TR><TD bgcolor=#FFFFFF class=calendary align=right>\n";
//document.selectDayTemplate += "		<A href=\"JavaScript:top.principal.updateDate('{_DAY_}-{_MONTH_}-{_YEAR_}')\" class=calendary>{_DAY_}</A></TD></TR></TABLE></TD>";
<%select case Action
	case 1
	%>
	document.selectDayTemplate += "		<A href=\"JavaScript:top.principal.agenda('{_DAY_}/{_MONTH_}/{_YEAR_}');top.principal.setDay('{_DAY_}/{_MONTH_}/{_YEAR_}')\" class=calendary>{_DAY_}</A></TD></TR></TABLE></TD>";
	<%
	case 2
	%>
  document.selectDayTemplate += "		<TR><TD bgcolor=#FFFFFF class=calendary align=right>{_DAYNUM_}</TD></TR></TABLE></TD>";	
	<%
	end select
	%>


</SCRIPT>
<SCRIPT language="javascript" src="javaScripts/calendaryLib.js"></SCRIPT>
<SCRIPT>
	function AgendarMensaje(DateSend){
					 top.opener.document.forms[0].<%=Label%>.value = DateSend;
					 top.close();
  }
	
	function agenda( DateSend ){
	var esp;
	<%select case Action
	case 1
	%>
	<%if Label = "DischargeDate" then%>
		if (top.opener.document.forms[0].DischargeDate.value != '') {com='\n';} else {com='';};
		top.opener.document.forms[0].DischargeDate.value = top.opener.document.forms[0].DischargeDate.value + com + DateSend;
	<%else%>
		top.opener.document.forms[0].<%=Label%>.value = DateSend;
	<%end if%>
		top.close();
	<%
	case 2
	%>
	document.forms[0].DateSend.value = DateSend;
	document.forms[0].DateOriginal.value = dateField;
	document.forms[0].FD.value = FirstDate;
	document.forms[0].Label.value = Label;
	document.forms[0].Subj.value = Subject;
	document.forms[0].Action.value = Action;
	document.forms[0].OID.value = ObjectID;
	document.forms[0].submit();
	<%
	end select
	%>
  }
	
  function updateDate( fecha, dia, mes, anio ){
	validateDate();
	var time = " "+document.clockFrame.document.time.hour.value+":"+document.clockFrame.document.time.minute.value+" "+document.clockFrame.document.time.meridian.value;
	alert(fecha+time);
	top.opener.document.forms[0].DateSend.value = fecha;
	//document.forms[0].DaySend.value = dia;
	//document.forms[0].MonthSend.value = mes;
	//document.forms[0].YearSend.value = anio;
	document.forms[0].submit;
	//top.opener.document.forms[0].fechaInicialBusqueda.value = fecha+time;
	//top.close();
  }
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="JavaScript:calendary(dateField);self.focus()">
	<TABLE border=0 cellspacing=0 cellpadding=2 width=100%>
	  <TR class=titlelist>
		<TD width=100% align=right>&nbsp;</TD>
	  </TR>
	</TABLE>
	<form action="Agenda_Hours.asp" name=forma method=post target=agenda>
				<input type=hidden value="" name=DateSend>
				<input type=hidden value="" name=FD>
				<input type=hidden value="" name=DateOriginal>
				<input type=hidden value="" name=Label>
				<input type=hidden value="" name=Subj>
				<input type=hidden value="" name=Action>
				<input type=hidden value="" name=OID>
	</form>	
</BODY>
</HTML>
<%end if%>
