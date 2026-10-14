
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


dim txtcari,fcba,txtsupplier ,panentext,spkno,job,section,fcbapost,tahuntanam,tiperkb
txtcari = Request("keyword")
fcba    = session("fcba")
panentext = Request("panen")
spkno   = Request("spkno")
veh     = Request("veh")
job     = Request("activity")
rkbno   = Request("rkbno")
tcrwt   = Request("tcrwt")
sectionname   = Request("sectionname")
tahun   = Request("tahun")
bulan   = Request("bulan")
fcbapost   = Request("fcba")
tahuntanam   = Request("tahuntanam")
tiperkb   = Request("tiperkb")



Set List = Server.CreateObject("ADODB.Recordset")

if panentext = "1" then 

    sqlview =  "select FCCODE as id, fcname || ' - ' || FCCODE as text from  ACTMASTER  " & _
    "  where fccode in('6502101','1321410') and fcname || FCCODE like UPPER('%" & txtcari & "%') order by FCNAME asc"

elseif panentext = 2 then 
   
    sqlview =  "select FCCODE as id, fcname || ' - ' || FCCODE as text from  ACTMASTER  " & _
    "  where fccode in('6310101','6210101') and fcname || FCCODE like UPPER('%" & txtcari & "%') order by FCNAME asc"
  
elseif panentext = 3 then 

    sqlview =  "SELECT DISTINCT co.jobcode id,c.fcname || '-' || c.fccode|| '- Rp.' || co.rate text " & _
    " FROM contractagreementdetail co " & _
    " INNER JOIN ACTMASTER c on (co.jobcode=c.FCCODE) " & _
    " where co.agreementcode='" & spkno & "' AND c.fcname || ' - ' || c.fccode like UPPER('%" & txtcari & "%') order by co.jobcode asc"
elseif panentext = 4 then 
	if (veh="WS1" or veh="WS-01") then 
    sqlview =  "select FCCODE as id, fcname || ' - ' || FCCODE as text from  JOBMASTER  " & _
    " where regexp_like(fccode,'^631|^641|^650|^131|^132|^661|^1130216|^1130204')  AND fccode <> '1114001'  AND fccode <> '6310101' AND fccode <> '6410101'  and fcname || FCCODE like UPPER('%" & txtcari & "%') order by FCNAME asc"
	else
	 sqlview =  "select FCCODE as id, fcname || ' - ' || FCCODE as text from  JOBMASTER  " & _
    "  where regexp_like(fccode,'^641|^650|^131|^132|^661|^1210301|^1130216|^1130204|^1210208') AND fccode <> '1114001' AND fccode <> '6410101'  and fcname || FCCODE like UPPER('%" & txtcari & "%') order by FCNAME asc"
	end if
elseif panentext = 5 then 
	
	sqlview =  "Select A.JOB AS ID,A.JOB || ' - ' || J.FCNAME AS TEXT " & _
    " From   ACTIVITY_NORM A " & _
    " LEFT JOIN JOBMASTER  J ON A.JOB=J.FCCODE " & _
    " WHERE A.FCBA='" & fcba & "' " & _
    " AND A.JOB || ' - ' || J.FCNAME like UPPER('%" & txtcari & "%') ORDER BY A.JOB ASC"

elseif panentext = 6 then 
    sqlview =  "Select luas_gross " & _
    " From FIELD " & _
    " WHERE FCBA='" & fcba & "'" & _ 
	" AND Fccode = '" & tcrwt & "'"
elseif panentext = 7 then 
    sqlview =  "select rate_1 from allowded where fcba ='" & fcba & "' and fccode ='A01' "
elseif panentext = 8 then 
    sqlview =  "Select A.MANDAYSNORM AS ID,A.MANDAYSNORM AS TEXT " & _
    " From   ACTIVITY_NORM A " & _
    " WHERE A.FCBA='" & fcba & "' AND A.JOB='" & job & "' " & _
    " AND A.MANDAYSNORM like UPPER('%" & txtcari & "%') ORDER BY A.JOB ASC"
elseif panentext = 9 then 
    sqlview =  "Select DISTINCT ACTCODE  AS ID,B.FCNAME AS TEXT " & _
    " From   RKBUPAHDTL A INNER JOIN JOBMASTER B ON a.ACTCODE=b.fccode " & _
    " Where rkbno='" & rkbno & "' AND  b.fcname || b.FCCODE like UPPER('%" & txtcari & "%') " & _ 
	" Order by FCNAME asc FETCH FIRST 100 ROWS ONLY"
    ' response.write sqlView
elseif panentext = 10 then 
    sqlview =  "Select uom_unit " & _
    " From JOBMASTER  " & _
    " Where FCCODE ='" & job & "' " 
elseif panentext = 11 then 
    sqlview =  "Select A.FCCODE AS ID,A.FCCODE || ' - ' || A.FCNAME AS TEXT " & _
    " From   jobmaster A " & _
    " WHERE A.FCCODE = '1321410' " & _
    " UNION " & _
    "Select A.FCCODE AS ID,A.FCCODE || ' - ' || A.FCNAME AS TEXT " & _
    " From   jobmaster A " & _
    " WHERE A.jobgroupcode in (select fccode from jobgroupmaster where activity_heading = 'PNN')  " & _
    " AND A.FCCODE || ' - ' || A.FCNAME like UPPER('%" & txtcari & "%') ORDER BY TEXT ASC"
elseif panentext = 12 then 
    sqlview =  "Select DISTINCT A.PLANTINGYEAR AS ID,A.PLANTINGYEAR AS TEXT " & _
    " From   ACTIVITY_NORM_HARV A " & _
    " WHERE A.FCBA='" & fcba & "' " & _
    " AND A.PLANTINGYEAR like UPPER('%" & txtcari & "%') ORDER BY A.PLANTINGYEAR ASC"
elseif panentext = 13 then 
    sqlview =  "Select rkbno " & _
    " From   rkbhdr " & _
    " WHERE fcba='" & fcba & "' " & _
    " AND sectioncode='" & sectionname & "' AND PERIOD_YEAR='" & tahun & "' AND PERIOD_MONTH='" & bulan & "' and DOCUMENT_CLASSIFICATION='" & tiperkb & "'"
elseif panentext = 14 then 
	
	sqlview =  "Select A.JOB AS ID,A.JOB || ' - ' || J.FCNAME AS TEXT " & _
    " From   MATERIALACTIVITY_NORM A " & _
    " LEFT JOIN JOBMASTER  J ON A.JOB=J.FCCODE " & _
    " WHERE A.FCBA='" & fcba & "' " & _
    " AND A.JOB || ' - ' || J.FCNAME like UPPER('%" & txtcari & "%') ORDER BY A.JOB ASC"
elseif panentext = 15 then 
	
	sqlview =  "Select A.FCCODE ID, A.FCCODE || ' - ' || A.FCNAME TEXT " & _
    " From  ACTIVEMATERIAL A " & _
    " INNER JOIN MATERIALACTIVITY_NORM M ON A.FCCODE=M.ACTMAT " & _
    " WHERE M.FCBA='" & fcbapost & "' AND M.JOB='" & job & "' " & _
    " AND A.FCCODE || ' - ' || A.FCNAME like UPPER('%" & txtcari & "%') ORDER BY A.FCCODE ASC"

elseif panentext = 16 then 
	
	sqlview =  "Select A.JOB AS ID,A.JOB || ' - ' || J.FCNAME AS TEXT " & _
    " From   VEHICLEACTIVITY_NORM A " & _
    " LEFT JOIN JOBMASTER  J ON A.JOB=J.FCCODE " & _
    " WHERE A.FCBA='" & fcba & "' " & _
    " AND A.JOB || ' - ' || J.FCNAME like UPPER('%" & txtcari & "%') ORDER BY A.JOB ASC"

elseif panentext = 17 then 
	
	sqlview =  "Select UOM " & _
    " From   VEHICLEACTIVITY_NORM " & _
    " WHERE FCBA='" & fcba & "' AND JOB='" & job & "' " & _
    " AND JOB like UPPER('%" & txtcari & "%') ORDER BY JOB ASC"

elseif panentext = 18 then 
	
	sqlview =  "Select jarak as id, jarak as text " & _
    " From   VEHICLEACTIVITY_NORM " & _
    " WHERE FCBA='" & fcba & "' AND JOB='" & job & "' " & _
    " AND JOB like UPPER('%" & txtcari & "%') ORDER BY jarak ASC"

elseif panentext = 19 then 
	
	sqlview =  "Select qty as id, qty as text " & _
    " From   VEHICLEACTIVITY_NORM " & _
    " WHERE FCBA='" & fcba & "' AND JOB='" & job & "' " & _
    " AND JOB like UPPER('%" & txtcari & "%') ORDER BY qty ASC"

elseif panentext = 20 then 
	
	sqlview =  "Select NORMA_HK, ROTASI " & _
    " From   ACTIVITY_NORM_HARV " & _
    " WHERE FCBA='" & fcba & "' AND PLANTINGYEAR='" & tahuntanam & "' "

else 

    sqlview =  "select FCCODE as id, fcname || ' - ' || FCCODE as text from  ACTMASTER  " & _
    "  where fcname || FCCODE like UPPER('%" & txtcari & "%') order by FCNAME asc FETCH FIRST 100 ROWS ONLY"


end if

'Response.Write sqlview

List.Open sqlview, connStr, 1
ListCount = List.RecordCount 

Set caljson = new JSON  'JSON.asp가 필요함
Set rs = connStr.Execute(sqlview)

Dim List_jsonString
List_jsonString = caljson.toJSON("results",rs, false)

Dim jsonObj, outputObj
set jsonObj = new JSONobject
set outputObj = jsonObj.parse(List_jsonString)  'json.objectClass.asp가 필요함(파싱하려면)


Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
oJSON.loadJSON(List_jsonString)

Response.write List_jsonString
%>