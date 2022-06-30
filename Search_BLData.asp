<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim GroupID, CountryDes, rs, Conn, CountTransit, Consolidated, aListValues, CountListValues, i, BLType, ConsignerID, DiceContenerTemp
	GroupID = CheckNum(Request("GID"))
    DiceContenerTemp = Request("DiceContenerTemp")
%>
 
<HTML><HEAD><TITLE>Aimar - Terrestre</TITLE></HEAD>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/config.js"></SCRIPT>
<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>
<SCRIPT language="javascript" src="img/validaciones.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript">
	function validate(){
	 	document.forma.submit();
	}	 	 
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="JavaScript:self.focus()">
	<FORM name="forma" action="Search_ResultsBLData.asp" method="post" target=_self>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<br>
	<TABLE cellspacing=0 cellpadding=2 width=100% align=center>
		<%select case GroupID
		Case 2, 33%>
		<TR>
		<TD class=label align=center colspan="2"><b>Agente</b></TD>
		</TR> 
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Name" value="" size=30></TD>
		</TR> 
		<% Case 3, 20%>
		<TR>
		<TD class=label align=center colspan="2"><b>Embarcador</b></TD>
		</TR> 
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Name" value="" size=30></TD>
		</TR> 
		<TR>
		<TD class=label align=right><b>Pais:</TD>
		<TD class=label align=left>
			<select name="Countries" id="Pais" class="label">
				<option value="">Seleccionar (Opcional)</option>
				<!--#include file=Countries.asp--> 
			</select>	
		</TD>
		</TR>
		<% Case 4, 11, 31, 34, 35, 36 'Consignatarios / Clientes / Coloader
			BLType = CheckNum(Request("BTP"))
		%>
		<TR>
		<TD class=label align=center colspan="2"><b>Consignatario / Exportador / Coloader</b></TD>
		</TR> 
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Name" value="<%=DiceContenerTemp%>" size=30></TD>
		</TR> 
		<TR>
		<TD class=label align=right><b>Pais:</TD>
		<TD class=label align=left>
			<input type="hidden" name="BTP" value="<%=BLType%>">
			<select name="Countries" id="Pais" class="label">
				<option value="">Seleccionar (Opcional)</option>
				<!--#include file=Countries.asp--> 
			</select>	
		</TD>
		</TR>
		<% Case 9 'Commodities
			BLType = CheckNum(Request("BTP"))
			CountryDes = Request("CTD")
			Consolidated = CheckNum(Request("CSD"))
			ConsignerID = CheckNum(Request("CID"))
			CountListValues = -1
		
			OpenConn Conn
				if Consolidated = 1 then
					Set rs = Conn.Execute("select a.BLDetailID from BLDetail a, BLs b where b.BLArrivalDate<>'' and a.InTransit=2 and b.Consolidated=1 and a.BLID=b.BLID and b.BLType=" & BLType & " and b.CountryDes in " & Session("Countries"))
				else
					Set rs = Conn.Execute("select count(a.BLDetailID), b.BLNumber from BLDetail a, BLs b where b.BLArrivalDate<>'' and a.InTransit=2 and b.Consolidated=0 and a.BLID=b.BLID and b.BLType=" & BLType & " and b.CountryDes in " & Session("Countries") & " and b.ConsignerID=" & ConsignerID & " group by b.BLNumber")
				end if
			'	1 = Actualizado, 2=Va en Transito, 3=Llego a Destino Final, 4=Itinerario Pendiente, 5=Itinerario Asignado
				if Not rs.EOF then
					aListValues = rs.GetRows
					CountListValues = rs.RecordCount-1
				end if	
			CloseOBJs rs, Conn
		 %>
		<TR>
		<TD class=label align=center colspan="2"><b>Productos</b></TD>
		</TR>
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="NameES" value="<%=DiceContenerTemp%>" size=30>
		<input type="hidden" name="CTD" value="<%=CountryDes%>">
		<input type="hidden" name="CID" value="<%=ConsignerID%>">
		<input type="hidden" name="CSD" value="<%=Consolidated%>">
		<input type="hidden" name="BTP" value="<%=BLType%>">
		<input type="hidden" name="BLN" value="">
		</TD>
		</TR> 
        <%
		if Consolidated = 1 then
			if CountListValues >= 0 then
		%>			
			<TR>
			<TD align="center" colspan="2">
			<TABLE cellspacing=0 cellpadding=4 width=400>
				<TR>
					<TD class=titlelist align="center">				
					<a href="#" onClick="Javascript:document.forma.GID.value=14;document.forma.submit();"><font color="FFFFFF"><b>Hay <% if CountListValues=1 then response.write CountListValues & " Producto" else  response.write CountListValues & " Productos" end if%> en Tr&aacute;nsito Consolidado</b></font></a>
					</TD>
				</TR>
				</TABLE>
			</TD>
			</TR>		
		<%
			end if
		else			
			for i=0 to CountListValues
		%>
			<TR>
			<TD align="center" colspan="2">
			<TABLE cellspacing=0 cellpadding=4 width=400>
				<TR>
					<TD class=titlelist align="center">				
					<a href="#" onClick="Javascript:document.forma.GID.value=14;document.forma.BLN.value='<%=aListValues(1,i)%>';document.forma.submit();"><font color="FFFFFF"><b>Hay <% if CInt(aListValues(0,i))=1 then response.write aListValues(0,i) & " Producto" else  response.write aListValues(0,i) & " Productos" end if%> en Tr&aacute;nsito Express de CP: <%=aListValues(1,i)%></b></font></a>
					</TD>
				</TR>
				</TABLE>
			</TD>
			</TR>		
		<%
			next
		end if
		Case 28 'Carga General
		%>
		<TR>
		<TD class=label align=center colspan="2"><b>Carga General</b></TD>
		</TR> 
		<TR>
		<TD class=label align=right width=40%><b>Tipo de Carga:</b></TD>
		<TD class=label align=left width=60%>
		<select name="BLType" class=label><!--Los codigos 4,5,6 y 7 son los generados en la base Master-->
			<%if Request("IT")=1 then%>
			<option value=4>TERRESTRE CONSOLIDADO</option>
			<option value=5>TERRESTRE EXPRESS</option>
			<%else%>
			<option value=6>TERRESTRE LOCAL</option>
			<!--<option value=7>TERRESTRE ENTREGA</option>-->
			<%end if%>
		</select>
		</TD>
		</TR>
		<TR>
		<TD class=label align=right><b>RO:</b></TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="RO" value="" size=30></TD>
		</TR> 			
		<TR>
		<TD class=label align=right><b>Pais:</b></TD>
		<TD class=label align=left>
		<select name="CountriesSearch" class=label>
			<!-- <option value=''>Seleccionar</option> 2021-11-23 al permitir que no seleccionen pais, al momento de de almacenar los valores de paises, si no lleva pais le quita los tres ultimos caracteres y pierde la empresa -->
			<%DisplayCountries Request("CTR"), 2%>
		</select>
		<input type="hidden" name="Countries" value="<%=Request("CTR")%>">
		</TD>
		</TR> 
		<% Case 29 'Rubros %>
		<TR>
		<TD class=label align=center colspan="2"><b>Rubros</b></TD>
		</TR> 
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left>
		<INPUT TYPE=text class=label name="Name" value="" size=30>
		<INPUT name="N" type=hidden value="<%=Request("N")%>">
		</TD>
		</TR> 
		<% Case 30 'Proveedores %>
		<TR>
			<%select Case Request("ST")%>
			<%case 0%>
			<TD class=label align=center colspan="2"><b>Lineas Aereas</b></TD>
			<%case 1%>
			<TD class=label align=center colspan="2"><b>Agentes</b></TD>
			<%case 2%>
			<TD class=label align=center colspan="2"><b>Navieras</b></TD>
			<%case 3%>		
			<TD class=label align=center colspan="2"><b>Proveedores</b></TD>
			<%end select%>
		</TR> 
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left>
		<INPUT TYPE=text class=label name="Name" value="" size=30>
		<INPUT name="N" type=hidden value="<%=Request("N")%>">
		<INPUT name="ST" type=hidden value="<%=Request("ST")%>">
		</TD>
		</TR> 		
		<%end select%>
		</TABLE>
		<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
		<TR>
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
			<TR>
				<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validate();" value="&nbsp;&nbsp;Buscar&nbsp;&nbsp;" class=label></TD>
			</TR>
			</TABLE>
		<TD>
		</TR>
	</TABLE>
	</FORM>
</BODY>

</HTML>
<%
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>
