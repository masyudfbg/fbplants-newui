<% Response.Buffer = False %>
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


dim txtcari,fcba,bkmfrom,bkmto,idview,company,filtersection,filterrowstate
idview   = Request("idview")
'tampilkan data listing bkm'
if idview = "1" then
    prcode = Replace(Request("prcode"), "'", "''") ' minimal sanitasi

    sqlview = "UPDATE PURCHASEREQUEST SET ROWSTATEID = 'Cancelled' WHERE PRCODE = '"&prcode&"'"
    connStr.Execute sqlview

    ' response.write sqlView

    if Err.Number <> 0 then
        Response.Write "{""status"":""error"",""message"":""Gagal mengubah data: " & Err.Description & """}"
    else
        Response.Write "{""status"":""success"",""message"":""Data berhasil diubah""}"
    end if
end if
%>