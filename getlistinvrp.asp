
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


dim txtcari,fcba,txtsupplier 
txtcari = Request("keyword")
txtnoinvoice    = Request("noinvoice")
txtjnsinvoice   = Request("jnsinvoice")
fcba = Session("fcBA")

Set List = Server.CreateObject("ADODB.Recordset")

'if txtjnsinvoice = "1" then

    sqlview =  "select h.totalamount" & _
	" from invoice h " & _
	" where h.invoicecode='" & txtnoinvoice & "'"

'else 
   '' sqlview =  "select h.grandtotal as totalamount " & _
''	" from invoiceservicehdr h " & _
''	" where h.invoiceno='" & selectnoinv & "'  " & _
 ''   " and h.invoiceno like UPPER('%" & txtcari & "%') order by h.invoiceno asc"

'end if

List.Open sqlview, connStr, 1
ListCount = List.RecordCount 

Set caljson = new JSON  'JSON.asp가 필요함
Set rs = connStr.Execute(sqlview)

Dim List_jsonString
List_jsonString = caljson.toJSON("datalist",rs, false)

Dim jsonObj, outputObj
set jsonObj = new JSONobject
set outputObj = jsonObj.parse(List_jsonString)  'json.objectClass.asp가 필요함(파싱하려면)


Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
oJSON.loadJSON(List_jsonString)

Response.write List_jsonString
%>