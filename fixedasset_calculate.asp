
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


dim txtcari,company,fcba,tahun,bulan,msg,txtperiod,prd,prd2
Dim TimeAwal,TimeAkhir,Hitung,Lama
TimeAwal=Time()
    company  = Request("company")
    fcba     = Request("fcba")
	tahun    = Request("tahun")
    bulan    = Request("bulan")
	
if trim(fcba) = "-" then 
	msg =  "<font face=Verdana size=1 color=red>" & "Silakan pilih Business Unit ! "
else	
	'-- Check for Empallded PreProcess
qCheckStatus = "SELECT Nvl(Max(To_Char(date_transaction,'YYYYMM')),'000000') period FROM Fixedasset_Depreciation_Ledger WHERE fcba='" & fcba & "'"
Set rsCheckStatus = Server.CreateObject("ADODB.Recordset")
rsCheckStatus.CursorType = 3
rsCheckStatus.CursorLocation = 3
rsCheckStatus.Open qCheckStatus, connstr
if not rsCheckStatus.eof then
   txtperiod = rsCheckStatus("period")
else
   msg =  "<font face=Verdana size=1 color=red>" & "period are not defined "
   txtperiod="000000"
end if
rsCheckStatus.close
set rsCheckStatus = Nothing

if txtperiod="000000" then 
	prd = cdbl(tahun)*100+cdbl(bulan)
else
    if right(txtperiod,2) = "12" then
	prd = cdbl(txtperiod) + 89
	else
	prd = cdbl(txtperiod) + 1
	end if
end if 
    prd2 = cdbl(tahun)*100+cdbl(bulan)

'response.write "period" & prd & "<br>"
'response.write "period2" & prd2


IF prd = prd2 THEN
	
	sqlexec = "BEGIN " & _ 
              "fac_journal_fac_fix_api.ProsesGenerate('" & tahun & "','" & bulan & "','" & fcba & "','system');" & _
			  "END; " 
    connStr.Execute(sqlexec)
  

    TimeAkhir=Time()
    Lama =TimeAkhir-TimeAwal

    msg =  "Proses Finish ... " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."
   
   
ELSE
    msg = "<font face=Verdana size=1 color=red>" & "Please proses period " & prd & " first.... !"
END IF	
end IF

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