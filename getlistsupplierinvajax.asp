
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


dim txtcari,fcba,jnsinvoice 
txtcari 	= Request("keyword")
jnsinvoice 	= Request("jnsinvoice")
fcba 		= Session("fcBA")

Set List = Server.CreateObject("ADODB.Recordset")

if jnsinvoice = "2" then

	sqlview =  " select c.customercode as id,c.description as text " & _
	" from customer c,invoiceservicehdr h " & _
	" where h.suppliername=c.customercode  " & _
	" and h.fcba=c.fcba and c.isactive='Y' AND c.supplier='Y' AND c.fcba=h.fcba " & _
	" AND h.grandtotal-h.payment>0 and h.invoice_by='" & fcba & "' " & _
	" and c.description like UPPER('%" & txtcari & "%') group by h.fcba,c.customercode,c.description " & _
	" order by c.description asc"
end if

if jnsinvoice = "1" then
	
	sqlview =  " select s.fccode as id,s.fcname as text " & _
	" from supplier s,invoice i " & _
	" where s.fccode=i.suppliercode  " & _
	" and s.fcba=i.fcba and s.activation='Y' " & _
	" AND i.totalamount-i.payment>0 and i.invoice_by='" & fcba & "' " & _
	" and s.fcname like UPPER('%" & txtcari & "%') " & _
	" group by s.fccode,s.fcname " & _
	" order by s.fcname asc"
end if
if jnsinvoice = "3" then 

	sqlview =  " select k.fccode as id,k.fcname as text  " & _
" from contractor k,contractorinvoicehdr i  " & _
" where k.fccode=i.contractorcode   " & _
" and k.fcba=i.fcba and k.activation='Y'  " & _
" AND i.grandtotal_aftertax-i.payment>0 and i.invoice_by='" & fcba & "' and k.fcname like UPPER('%" & txtcari & "%')  group by k.fccode,k.fcname  " & _
"  UNION ALL " & _
" select k.fccode as id,k.fcname as text  " & _
" from contractor k,CONTRACTORINVOICEDP  j " & _
" where k.fccode=j.contractorcode   " & _
" and k.fcba=j.fcba and k.activation='Y' and j.invoice_by='" & fcba & "'  " & _
" and j.dp-j.payment>0 and k.fcname like UPPER('%" & txtcari & "%') " & _
" group by k.fccode,k.fcname"
 
 end if
	 
if jnsinvoice = "4" then 

	sqlview =  "select k.fccode as id,k.fcname as text " & _
	" from supplier k,invoicedphdr i " & _
	" where k.fccode=i.suppliercode  " & _
	" and k.fcba=i.fcba  " & _
	" AND i.totalamount-i.payment>0 and i.invoice_by='" & fcba & "' " & _
	" and k.fcname like UPPER('%" & txtcari & "%') " & _
	" group by k.fccode,k.fcname " & _
	" order by k.fcname asc"
 
 end if
if jnsinvoice = "5" then 

	sqlview =  "select k.fccode as id,k.fcname as text " & _
	" from transporter k,invoicesto i " & _
	" where k.fccode=i.transportercode  " & _
	" and k.fcba=i.fcba  " & _
	" AND i.totalamount-i.payment>0 and i.invoice_by='" & fcba & "' " & _
	" and k.fcname like UPPER('%" & txtcari & "%') " & _
	" group by k.fccode,k.fcname " & _
	" order by k.fcname asc"
 
 end if

 if jnsinvoice = "6" then

	sqlview = "SELECT s.fccode as id,s.fcname as text " & _
	 " from supplier s,invoiceservicehdr h  " & _
	 " where s.fccode=h.suppliername and s.fcba=h.fcba  " & _
	 " and s.ACTIVATION='Y' AND  h.grandtotal-h.payment>0 and h.invoice_by='" & fcba & "'  " & _
	 " AND s.fcname LIKE UPPER('%" & txtcari & "%') " & _
	 " group by h.fcba,s.fccode,s.fcname " & _
	 " ORDER BY s.fccode,s.fcname ASC"
	
 end if

 if jnsinvoice = "7" then

	sqlview = "SELECT t.fccode as id,t.fcname as text " & _
	" from TRANSPORTER t,TRANSPORTERCONTRACT tc,TRANSPORTERINVOICE ti " & _
	" where t.FCCODE=tc.TRANSPORTER and ti.TRANSPORTERCONTRACTNO=tc.TRANSPORTERCONTRACTNO " & _
	" AND t.fcba=tc.FCBA and ti.fcba=tc.fcba " & _
	" and ti.invoice_by='" & fcba & "' " & _
	" AND ti.AMOUNTAFTERPPH-ti.payment>0 AND t.fcname LIKE UPPER('%" & txtcari & "%') " & _
	" group by t.fccode,t.fcname " & _
	" order by t.fcname asc"

 end if
 if jnsinvoice = "10" then

	sqlview = "SELECT distinct t.fcba as id,t.fcba as text " & _
	" from OPR_INV_PDO t" & _
	" where t.invoiceby='" & fcba & "' " & _
	" AND t.amount-t.payment>0 AND t.fcba LIKE UPPER('%" & txtcari & "%') " & _
	" order by t.fcba asc"

 end if

 ' get supplier tbs
if jnsinvoice = "100" then 
	if (fcba="ARM") then 
	sqlview =  "SELECT CUSTOMERCODE AS id,CUSTOMERCODE || ' - ' || DESCRIPTION AS text" & _
	" from CUSTOMER " & _
	" where fcba='ARM' AND ISACTIVE = 'Y'" & _
	" and CUSTOMERCODE || ' - ' || DESCRIPTION like UPPER('%" & txtcari & "%') " & _
	" UNION ALL " & _
	" SELECT FCCODE AS id,FCCODE || ' - ' || FCNAME AS text " & _
	" FROM SUPPLIER " & _
	" WHERE FCTYPE = 'STY02' AND ACTIVATION = 'Y' And fcba='ARM'"
	else
	sqlview =  "SELECT CUSTOMERCODE AS id,CUSTOMERCODE || ' - ' || DESCRIPTION AS text" & _
	" from CUSTOMER " & _
	" where fcba='" & fcba & "' AND ISACTIVE = 'Y'" 
	end if
 end if

  ' get product
if jnsinvoice = "200" then 

sqlview =  "select PRODUCTCODE as id,PRODUCTCODE || ' - ' || PRODUCTDESCRIPTION as text " & _
" from PRODUCT " & _
" where fcba='" & fcba & "' " & _
" and PRODUCTCODE || ' - ' || PRODUCTDESCRIPTION like UPPER('%" & txtcari & "%') " & _
" order by PRODUCTDESCRIPTION asc"

end if

  ' get customer
if jnsinvoice = "8" then 

sqlview =  "SELECT distinct " & _ 
" sih.customercode as id, " & _
" sih.customercode || ' - ' || customer_api.get_description (sih.customercode, sih.fcba) as text " & _
" FROM salesinvoicehdr sih " & _
" INNER JOIN salesagreementdet sah " & _
" ON (    sah.fcba = sih.fcba AND sah.salescontractno = sih.salescontractno)  " & _
" where (sih.invoice_by='" & fcba & "' OR sih.invoice_by='" & fcba & "') AND sih.grandtotal-sih.RECEIVEAMT>0"

end if

if jnsinvoice = "9" then
	
	sqlview =  " select s.fccode as id,s.fcname as text " & _
	" from supplier s,purchaseinvoicehdr i " & _
	" where s.fccode=i.suppliercode  " & _
	" and s.fcba=i.fcba and s.activation='Y' " & _
	" and i.invoice_by='" & fcba & "' " & _
	" and s.fcname like UPPER('%" & txtcari & "%') " & _
	" group by s.fccode,s.fcname " & _
	" order by s.fcname asc"
end if	 


List.Open sqlview, connStr, 1
ListCount = List.RecordCount 

Set caljson = new JSON  'JSON.asp가 필요함
Set rs = connStr.Execute(sqlview)

Dim List_jsonString
List_jsonString = caljson.toJSON("results", rs, false)

Dim jsonObj, outputObj
set jsonObj = new JSONobject
set outputObj = jsonObj.parse(List_jsonString)  'json.objectClass.asp가 필요함(파싱하려면)


Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
oJSON.loadJSON(List_jsonString)

Response.write List_jsonString
%>