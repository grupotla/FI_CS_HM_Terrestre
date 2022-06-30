<%@ Language=VBScript %>
<!-- #INCLUDE file="utils.asp" -->
<%
Response.write Request.ServerVariables("REMOTE_ADDR") & "<br>"
Response.write Request.ServerVariables("GATEWAY_INTERFACE") & "<br>"
Response.write Request.ServerVariables("REMOTE_HOST") & "<br>"
Response.write Request.ServerVariables("REQUEST_METHOD") & "<br>"
Response.write Request.ServerVariables("URL") & "<br>" 
%>
