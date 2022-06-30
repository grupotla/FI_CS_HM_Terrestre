<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"

Dim ObjectID, esquema, Conn, rs, Action, QuerySelect, QuerySelect2, QuerySelect3, i, Asigns, Ides, cagen, ctrans, cotros, j, Pos, k, l, Res, Homologado, ValidoHomo, AgentTypeItemID
'Dim Freight, Freight2, Insurance, Insurance2, AnotherChargesCollect, AnotherChargesPrepaid

ObjectID = CheckNum(Request("OID"))
Action = Request("Action")
Pos = Request("Pos")
esquema = Request("esquema")

if Action <> "" then

    OpenConn Conn

    Response.Write "ENTRO (" & Action & ")(" & ObjectID & ")(" & esquema & ")<br>"
%>

<body onload="refresh();">

</body>

<script type="text/javascript">
    function refresh() {
        parent.document.forma.Action.value = 0;
        parent.document.forma.target = "";
        parent.document.forma.action = "ItineraryCharges.asp";

        window.parent.document.location.reload(true);

        //window.onbeforeunload = null;
        //parent.location.reload();

    }
</script>

<%
    'if Action=2 or Action=994 or Action=995 then '994 update    995 insert
		
    '    SaveChargeItems Conn, ObjectID, Action, 0
            
	'end if

    if Action = "asignar" or Action = "liberar" then 'asignar clientes

        Response.Write Request.Form("CHK") & "<br>"

        Asigns=split(Request.Form("CHK"),",") 'siempre viene solo un registro NO puede traer los que el usuario seleccione

        For i=LBound(Asigns) to UBound(Asigns)    
                
            Ides=split(Asigns(i),"|") '0 ChargeID, 1 id_pedido, 2 pedido_erp, 5 cantidad rubros con pedido_id 

            select case Action 
                case "998", "994" 'delete / update

                    QuerySelect = "UPDATE ChargeItems SET Expired = 1 WHERE ChargeID = '" & Ides(0) & "'"
                    Response.Write QuerySelect & "<br>"
                    Conn.Execute(QuerySelect)
                
                    SaveChargeBL Conn, ObjectID

                case "liberar" 'liberar pedido_id   '997 'liberar pedido_erp   id_cliente = NULL, cliente_nombre = NULL, 
                
                    QuerySelect = "UPDATE ChargeItems SET id_pedido = NULL, pedido_erp = NULL, InvoiceID = 0, DocType = 0 WHERE ChargeID = '" & Ides(0) & "'"  
                    Response.Write QuerySelect & "<br>"
                    Conn.Execute(QuerySelect)
                
                'case 994, 995, 999 '994 update    995 insert    999 vincular cliente a rubro existente
                
                case "asignar" 'solo cuando asigna cliente

                    Res = ValidaHomologacion("1", esquema, "01", "'" & Ides(3) & "'")   

                    On Error Resume Next
                        Homologado = IFNULL(Res(0,0))
                    If Err.Number <> 0 Then
                        Homologado = IFNULL(Res)
                    end if

                    'Response.Write "(Homologado=" & Homologado & ")<br>"

                    if Homologado <> "" then
                    
                        j = UBound(Res)

                        'Response.Write "(j=" & j & ")<br>"
                        
                        'if j > 0 then

                            'Response.Write "(" & Res(0,0) & ")<br>"
                            'Response.Write "(" & Res(1,0) & ")<br>"
                            'Response.Write "(" & Res(2,0) & ")<br>"
                            'Response.Write "(" & Res(3,0) & ")<br>"
    
                        'end if

                        QuerySelect = "UPDATE ChargeItems SET id_cliente = " & Ides(3) & ", cliente_nombre = '" & Ides(4) & "', id_pedido = NULL, pedido_erp = NULL, InvoiceID = 0, DocType = 0 WHERE ChargeID = '" & Ides(0) & "'"
                        Response.Write QuerySelect & "<br>"
                        Conn.Execute(QuerySelect)


                    else

                        response.write "<script" & ">alert('Cliente seleccionado no esta homologado');</script>"
    
                    end if

            end select 

        Next 

        CloseOBJ Conn
    
        Response.Write "Finalizo Proceso<br>"

        Response.End

	end if
   




'/////////////////////////////////////////////////// SECCION RUBROS 

    Dim CantItems, CreatedDate, CreatedTime, ItemCurrs, ItemIDs, ItemVals, ItemLocs, ItemNames, ItemOVals, ItemPPCCs, ItemServIDs, ItemChargeID, ItemPedErp, ItemServNames, ItemInvoices, ItemCalcInBls, ItemInRO, CType, ItemInterCompanyIDs, ItemDocType, ItemCli, ItemCliNom, ItemRegimen, ItemTarifaPrice, ItemTarifaTipo, ItemAgent

	FormatTime CreatedDate, CreatedTime
	
	CantItems = CheckNum(Request.Form("CantItems"))
	ItemCurrs = Split(Request.Form("ItemCurrs"), "|")
	ItemIDs = Split(Request.Form("ItemIDs"), "|")
	ItemServIDs = Split(Request.Form("ItemServIDs"), "|")
	ItemServNames = Split(Request.Form("ItemServNames"), "|")
	ItemVals = Split(Request.Form("ItemVals"), "|")
	ItemLocs = Split(Request.Form("ItemLocs"), "|")
	ItemNames = Split(Request.Form("ItemNames"), "|")
	ItemOVals = Split(Request.Form("ItemOVals"), "|")
	ItemPPCCs = Split(Request.Form("ItemPPCCs"), "|")
	ItemInvoices = Split(Request.Form("ItemInvoices"), "|")
	ItemDocType = Split(Request.Form("ItemDocType"), "|")
	ItemCalcInBLs = Split(Request.Form("ItemCalcInBls"), "|")
	ItemInRO = Split(Request.Form("ItemInRO"), "|")
	ItemInterCompanyIDs = Split(Request.Form("ItemInterCompanyIDs"), "|")                
    ItemChargeID = Split(Request.Form("ItemChargeID"), "|")
    ItemAgent = Split(Request.Form("ItemAgent"), "|")

    if Request.Form("ItemCli") = "" then
        ItemCli = Split("0", "|")    
    else
        ItemCli = Split(Request.Form("ItemCli"), "|")
    end if

    if Request.Form("ItemPedErp") = "" then
        ItemPedErp = Split(" ", "|")    
    else
        ItemPedErp = Split(Request.Form("ItemPedErp"), "|")
    end if

    if Request.Form("ItemCliNom") = "" then
        ItemCliNom = Split(" ", "|")    
    else    
        ItemCliNom = Split(Request.Form("ItemCliNom"), "|")
    end if

    if Request.Form("ItemRegimen") = "" then
        ItemRegimen = Split(" ", "|")    
    else    
        ItemRegimen = Split(Request.Form("ItemRegimen"), "|")
    end if

    if Request.Form("ItemTarifaPrice") = "" then
        ItemTarifaPrice = Split(" ", "|")    
    else    
        ItemTarifaPrice = Split(Request.Form("ItemTarifaPrice"), "|")
    end if

    if Request.Form("ItemTarifaTipo") = "" then
        ItemTarifaTipo = Split(" ", "|")    
    else    
        ItemTarifaTipo = Split(Request.Form("ItemTarifaTipo"), "|")
    end if



    select case Action 
        
        case "insert", "update", "borrar"

            cagen = -1
            ctrans = -1
            cotros = -1

            ValidoHomo = true
            QuerySelect3 = ""
            QuerySelect = ""

            for i=0 to CantItems

                if Action = "borrar" and Cdbl(Pos) = Cdbl(i) then
                    'cuando borra no acumula el rubro a la guia

                else

                    k = 0
                    l = 0

                end if

                'Homologado = "*" 'debe llevar algo

                'Response.Write "ItemAgent=" & ItemAgent(i) & " Agen=" & cagen & " Tran=" & ctrans & " Otr=" & cotros & "<br>"

                Response.Write "Pos=" & Pos & " i=" & i & "<br>"


                if Cdbl(Pos) = Cdbl(i) then


                    if Action = "insert" then                        

                        Res = ValidaHomologacion("1", esquema, "01", "'" &  CheckNum(ItemCli(i)) & "'")   

                        On Error Resume Next
                            Homologado = IFNULL(Res(0,0)) 'si trae valor 
                            ValidoHomo = true
                        If Err.Number <> 0 Then
                            Homologado = IFNULL(Res) 'aca asigna blancos
                            ValidoHomo = false
                        end if

                    end if

                    Response.Write "(Homologado=" & ValidoHomo & ")<br>"

                    if Action = "update" or Action = "borrar" then                        
                        'QuerySelect2 = "UPDATE ChargeItems SET Expired=1 WHERE ChargeID = " & ItemChargeID(i) 

                        QuerySelect2 = "UPDATE ChargeItems SET Expired=1 WHERE SBLID = " & ObjectID & " AND ItemID = " & CheckNum(ItemIDs(i))
                        Response.Write QuerySelect2 & "<br>"
                        Conn.Execute(QuerySelect2)
                    end if

        'Response.Write "Pos=" & Pos & " i=" & i & "<br>"
                
        'Response.Write "1*ENTRO (" & ItemCurrs(i) & ")(" & ItemIDs(i) & ")(" & ItemVals(i) & ")(" & ItemLocs(i) & ")(" & ItemAgent(i) & ")<br>"

        'Response.Write "2*ENTRO (" & ItemNames(i) & ")(" & ItemServIDs(i) & ")(" & ItemServNames(i) & ")(" & ItemPPCCs(i) & ")(" & ItemTarifaPrice(i) & ")(" & ItemRegimen(i) & ")(" & ItemTarifaTipo(i) & ")<br>"

        'Response.Write "3*ENTRO (" & ItemCli(i) & ")(" & ItemPedErp(i) & ")(" & ItemCliNom(i) & ")<br>"

                    if Action = "update" or (Action = "insert" and ValidoHomo = true) then                        
                
                        QuerySelect2 = "INSERT INTO ChargeItems (SBLID, Currency, ItemID, Value, Local, AgentTyp, ItemName, CreatedDate, CreatedTime, UserID, ServiceID, ServiceName, PrepaidCollect, CalcInBL, TarifaPricing, Regimen, TarifaTipo, id_cliente, pedido_erp, cliente_nombre) VALUES " & _                    
                        "(" & ObjectID  & ", '" & ItemCurrs(i) & "', " & CheckNum(ItemIDs(i)) & ", " & CheckNum(ItemVals(i))  & ", " & CheckNum(ItemLocs(i))  & ", " & CheckNum(ItemAgent(i)) & ", " & _ 
                        "'" & ItemNames(i) & "', '" & CreatedDate & "', " & CreatedTime & ", " & Session("OperatorID") & ", " & CheckNum(ItemServIDs(i)) & ", '" & ItemServNames(i) & "', " & CheckNum(ItemPPCCs(i))  & ", 0, '" & CheckNum(ItemTarifaPrice(i)) & "', '" & ItemRegimen(i) & "', '" & ItemTarifaTipo(i) & "', " & CheckNum(ItemCli(i)) & ", '" & Trim(ItemPedErp(i)) & "', '" & Left(Trim(ItemCliNom(i)),99) & "')"    				
                        Response.Write QuerySelect2 & "<br>"
                        Conn.Execute(QuerySelect2)

                    end if


                else

                    if Action = "borrar" then                        

                        'QuerySelect2 = "UPDATE ChargeItems SET Pos=" & CheckNum(l) & " WHERE ChargeID = " & ItemChargeID(i) 
                        'Response.Write QuerySelect2 & "<br>"
                        'Conn.Execute(QuerySelect2)

                    end if

                end if

            next

                    'Response.Write "<br><br>" & QuerySelect & "<br><br>"


            if ValidoHomo = true then

                if QuerySelect3 <> "" then
                    'Response.Write QuerySelect3 & "<br>"
                    'Conn.Execute(QuerySelect3)
                end if

                if QuerySelect <> "" then
                    'QuerySelect = "UPDATE " & Iif(AWBType = "1","Awb","Awbi") & " SET Expired = 0" & QuerySelect & " WHERE SBLID = " & ObjectID
                    'Response.Write QuerySelect & "<br>"
                    'Conn.Execute(QuerySelect)
                end if


            else
                
                if Action = "insert" then
                    response.write "<script" & ">alert('Cliente seleccionado no esta homologado');</script>"
                end if

            end if

            Response.Write "Finalizo Proceso<br>"


    end select 

   CloseOBJ Conn


end if
%>


