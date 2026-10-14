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


Dim  txtcari,fcba,tgl,msg,tglSekarang,qtyHari,isHO,getHO
Dim txtClosingStatus,txtClosingAcc,txtClosingHari
    getHO = Request("getHO")
	if len(getHO)=0 or getHO="" then
		getHO = 0
	end if
	
	fcba = Request("fcba")
	if len(fcba)=0 or fcba="" then
		fcba = session("fcba")
	end if
	
	tgl   = Request("tglbkm")
 	tglSekarang = textDate(Now())
	qtyHari = GetQtyHariKerja(tgl,tglSekarang)
	
qCekHO = "SELECT fctype FROM businessunit where (fctype='HO' or fctype='R') and fccode='" & fcba & "' and 1=" & getHO &""
Set rsCekHO = Server.CreateObject("ADODB.Recordset")
rsCekHO.CursorType = 3
rsCekHO.CursorLocation = 3
rsCekHO.Open qCekHO, connstr
if not rsCekHO.eof  then 
	isHO = "TRUE"
else
    isHO = "FALSE"
end if
rsCekHO.close
set rsCekHO = Nothing

	if isHO = "FALSE" then
		if (qtyHari > 3 OR qtyHari < 0) then
			txtClosingHari = "TRUE"
		else
			txtClosingHari = "FALSE" 
		end if
	else
	   txtClosingHari = "FALSE"
	end if
	
	
	
'-- Check for Empallded PreProcess
qCheckStatus = " SELECT empalldedstatus_api.GetClosePayroll(to_date('"&tgl&"', 'mm/dd/yyyy'),'" & fcba & "') closingStatus, " & _ 
               " empalldedstatus_api.getstatusclosing_acc(EXTRACT(YEAR FROM to_date('"&tgl&"', 'mm/dd/yyyy')),EXTRACT(MONTH FROM to_date('"&tgl&"', 'mm/dd/yyyy')),'" & fcba & "') closingAcc " & _
               " FROM dual"
' response.write qCheckStatus

Set rsCheckStatus = Server.CreateObject("ADODB.Recordset")
rsCheckStatus.CursorType = 3
rsCheckStatus.CursorLocation = 3
rsCheckStatus.Open qCheckStatus, connstr
if not rsCheckStatus.eof then
   txtClosingStatus = rsCheckStatus("closingStatus")
   txtClosingAcc    = rsCheckStatus("closingAcc")
    IF txtClosingStatus = "FALSE" and txtClosingAcc = "FALSE" and txtClosingHari= "FALSE" then
		txtClosingStatus="FALSE"
	else
		'txtClosingStatus="FALSE"   
		txtClosingStatus="TRUE"
	END IF
else
   txtClosingStatus="TRUE"
end if
rsCheckStatus.close
set rsCheckStatus = Nothing
	

msg =  txtClosingStatus

Set List = Server.CreateObject("ADODB.Recordset")
sqlview = "select  '" & msg & "' as msg from dual"
    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonStringgl
    List_jsonStringgl = caljson.toJSON("results",rs, false)

    Dim jsonObjgl, outputObjgl
    set jsonObjgl = new JSONobject
    set outputObjgl = jsonObjgl.parse(List_jsonStringgl)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonStringgl)

    Response.write List_jsonStringgl

    
     
%>