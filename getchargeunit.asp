
<%
Response.CharSet="euc-kr"
Session.codepage="949"
Response.codepage="949"
Response.ContentType="text/html;charset=euc-kr"
%>
<%
Response.LCID = 1046 
%>

<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->

<!--#includes file="include/JSON.asp" --> 
<!--#includes file="include/json.objectClass.asp"-->
<!--#includes file="include/aspJSON.asp" -->

<%


dim txtcari,txtnoinvoice,jenisinv,sqlinvpuch
txtcari = Request("keyword")
txtnoinvoice = Request("noinvoice")
jenisinv = Request("jnsinvoice")

    Set List = Server.CreateObject("ADODB.Recordset")

   
	' get invoice purchasing
	IF jenisinv="1" THEN 

		sqlinvpuch = "select suppliercode AS chargecode,fcba as chargeunit" & _
		" from invoice " & _
		" where invoicecode='" & txtnoinvoice & "' "
		
	end if
	
	' get invoice tbs
	if jenisinv="2" THEN
		
		sqlinvpuch = "select SUPPLIERNAME AS chargecode,fcba as chargeunit" & _
		" from invoiceservicehdr " & _
		" where INVOICENO='" & txtnoinvoice & "' "
		
	end if
	' get invoice contractor
	if jenisinv="3" THEN
		
		sqlinvpuch = "select contractorcode AS chargecode,fcba as chargeunit " & _
		" from contractorinvoicehdr " & _
		" where INVOICECODE='" & txtnoinvoice & "'  " & _ 
		" 	union all " & _ 
		" select contractorcode AS chargetype,fcba as chargeunit " & _
		" from CONTRACTORINVOICEDP " & _
		" where INVOICECODE='" & txtnoinvoice & "'"  
		
	end if
	' get invoice payment request
	if jenisinv="4" THEN
		
		sqlinvpuch = "select suppliercode AS chargecode,fcba as chargeunit " & _
		" from INVOICEDPHDR " & _
		" where INVOICECODE='" & txtnoinvoice & "' "
		
	end if
	' get invoice STO
	if jenisinv="5" THEN
		
		sqlinvpuch = "select transportercode AS chargecode,fcba as chargeunit " & _
		" from INVOICESTO " & _
		" where INVOICECODE='" & txtnoinvoice & "' "
		
	end if
	' get SPESIAL INVOICE
	if jenisinv="6" THEN
		
		sqlinvpuch = "select suppliername AS chargecode,fcba as chargeunit " & _
		" from INVOICESERVICEHDR " & _
		" where INVOICENO='" & txtnoinvoice & "' "
		
	end if
	' get TRANSPORTER INVOICE
	if jenisinv="7" THEN
		
		sqlinvpuch = "SELECT tc.TRANSPORTER AS chargecode,tc.fcba as chargeunit " & _
		" from TRANSPORTER t,TRANSPORTERCONTRACT tc,TRANSPORTERINVOICE ti  " & _
		" where t.FCCODE=tc.TRANSPORTER and ti.TRANSPORTERCONTRACTNO=tc.TRANSPORTERCONTRACTNO " & _
		" AND t.fcba=tc.FCBA and ti.fcba=tc.fcba  " & _
		" and ti.INVOICECODE ='" & txtnoinvoice & "' " & _
		" group by tc.TRANSPORTER,tc.fcba"
		
	END IF


    List.Open sqlinvpuch, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlinvpuch)

    Dim List_jsonString
    List_jsonString = caljson.toJSON("dataview",rs, false)

    Dim jsonObj, outputObj
    set jsonObj = new JSONobject
    set outputObj = jsonObj.parse(List_jsonString)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString)

    Response.write List_jsonString


%>