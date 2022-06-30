<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"
Dim GroupID, BLID, SBLID, BLType, ClientID, AgentID, Sep, Countries, EXID

    GroupID = CheckNum(Request("GID"))
	BLID = CheckNum(Request("BLID"))
	SBLID = CheckNum(Request("SBLID"))
	BLType = CheckNum(Request("BTP"))
	ClientID = CheckNum(Request("CID"))
	AgentID = CheckNum(Request("AID"))
	Sep = CheckNum(Request("SEP"))
    Countries = Request("CTR")
    EXID = CheckNum(Request("id_routing"))
%>
<HTML>
<HEAD>
<SCRIPT language="javascript"> 
function imprimir()
{ if ((navigator.appName == "Netscape")) { window.print() ; 
} 
else
{ var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>'; 
document.body.insertAdjacentHTML('beforeEnd', WebBrowser); WebBrowser1.ExecWB(7, -1); WebBrowser1.outerHTML = "";
}
}
</SCRIPT> 
</HEAD>
<!--<BODY onload="imprimir();">-->
<BODY>


<%="" %>
	<iframe src="BLPrint.asp?GID=<%=GroupID%>&BLID=<%=BLID%>&SBLID=<%=SBLID%>&BTP=<%=BLType%>&CID=<%=ClientID%>&AID=<%=AgentID%>&SEP=<%=Sep%>&id_routing=<%=EXID%>" frameborder="0" framespacing="0" scrolling="auto" width="750" height="1500">
    
<%="" %>
    <iframe src="Conditions.asp?EXDBCountry=<%=Countries%>" frameborder="2" framespacing="0" scrolling="auto" width="750" height="960">
<%="paso3<br>" %>
    
</BODY>
</HTML>