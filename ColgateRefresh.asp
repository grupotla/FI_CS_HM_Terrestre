<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="Utils.asp" -->


<script type="text/javascript">
    
    setTimeout(function () {
        window.location.reload();
    }, 5000);
    
    //const d = new Date();
    //document.getElementById("container").innerHTML = dato + ' ' + d.toLocaleTimeString();

</script>

<style>

body { color:white; font-family:arial; font-size:10px}

@keyframes blink {
  50% {
    opacity: 0.0;
  }
}
.blink {
  animation: blink 1s step-start 0s infinite;
  font-size:13px;
}

</style>

<body>

<%

Dim Conn, rs, resultStr, QuerySelect

Dim dd, mm, yy, hh, nn, ss
Dim datevalue, timevalue, dtsnow, dtsvalue

'Store DateTimeStamp once.
dtsnow = Now()

'Individual date components
dd = Right("00" & Day(dtsnow), 2)
mm = Right("00" & Month(dtsnow), 2)
yy = Year(dtsnow)
hh = Right("00" & Hour(dtsnow), 2)
nn = Right("00" & Minute(dtsnow), 2)
ss = Right("00" & Second(dtsnow), 2)

'Build the date string in the format yyyy-mm-dd
datevalue = yy & "-" & mm & "-" & dd
'Build the time string in the format hh:mm:ss
timevalue = hh & ":" & nn & ":" & ss
'Concatenate both together to build the timestamp yyyy-mm-dd hh:mm:ss

    resultStr = timevalue

    On Error Resume Next

        if Session("PerfilColgate") = 1 then

            QuerySelect = GetHTMLSource("http://10.10.1.20/colgate/ColgateResponse.php?accion=Consumir")
            'QuerySelect = GetHTMLSource("http://localhost/colgate/ColgateResponse.php?accion=Consumir")

            'response.write QuerySelect & "<br>"

	        openConn2 Conn      

            QuerySelect = "SELECT COUNT(a.tch_pk) " & _

            "       FROM ti_colgate_header a " 


            QuerySelect = QuerySelect & "INNER JOIN ( " & _

            "    SELECT count(*), MAX(purposecode), MIN(purposecode), shipment, CASE WHEN BLDetailID in (0,-1) THEN MIN(purposecode) ELSE MAX(purposecode) END purposecode1 " & _
            "    FROM ( " & _
            "        SELECT a.""tch_b2_ShipmentIDNumber"" shipment, a.""tch_SetPurposeCode"" purposecode, " & _
            "        COALESCE((SELECT b.""tch_BLDetailID"" FROM ti_colgate_header b WHERE b.""tch_b2_ShipmentIDNumber"" = a.""tch_b2_ShipmentIDNumber"" AND b.""tch_SetPurposeCode"" <> '02' AND b.tch_estado = 2 ORDER BY b.""tch_SetPurposeCode"" DESC LIMIT 1),-1) as BLDetailID " & _ 
            "        FROM ti_colgate_header a  " & _
            "        WHERE a.tch_estado = 1  " & _
            "    ) x GROUP BY shipment, BLDetailID  " & _

            ") y ON shipment = a.""tch_b2_ShipmentIDNumber"" AND purposecode1 = a.""tch_SetPurposeCode"" " & _

            "WHERE a.tch_estado = 1"

            'response.write QuerySelect & "<br>"

	        Set rs = Conn.Execute(QuerySelect)
	        if Not rs.EOF Then

                resultStr = "Tienes " & "<span class=blink>(" & rs(0) & ")</span>" & " Registros Colgate Pendientes."  & " " & timevalue

	        end if

        end if

        response.write resultStr 

    If Err.Number<>0 then                                '
	    'response.write " Error : " & Err.Number & "  " & Err.Description & "<br>"
    end if
%>
</body>

