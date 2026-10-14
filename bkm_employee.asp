
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


dim txtcari,fcba,gangcode,upahedit,bkmno,tglbkm,ovtimehours,typeovrtime,jobcode,stocksubgroupcode,tahun,bulan,modulid
fcba = session("fcba")
txtcari = Request("keyword")
gangcode = Request("gangcode")
vehiclecode = Request("vehiclecode")
itemcode = Request("itemcode")
satuan = Request("satuan")
tglbkm = Request("tglbkm")
ovtimehours = Request("ovtimehours")
typeovrtime = Request("typeovrtime")
employee    = Request("employee")
sectionname    = Request("sectionname")
employeecode    = Request("emp")
jobcode    = Request("jobcode")
stocksubgroupcode    = Request("stocksubgroupcode")
tahun    = Request("tahun")
bulan    = Request("bulan")
modulid    = Request("modulid")


if satuan =1 then 
    sqlview = " select unitofmeasurement from stock where fccode='" & itemcode & "' and fcba='" & fcba & "'"
elseif satuan = 7  then 
    sqlview = " select nvl(convertion,0) convertion from overtime_convertion where type = '" & typeovrtime & "' and hours = '" & ovtimehours & "' and fcba = '" & fcba & "'"
elseif satuan = 8  then 
    sqlview = " select uomactivematerial,actmatnorm from materialactivity_norm where actmat='" & itemcode & "' and fcba='" & fcba & "' AND job='" & jobcode & "'"
else 

'AMBIL EMPLOYE DARI INPUT RAWATAN 
if employee = 100 then

 sqlview =  "select bkmno, sectioncode,gangcode,bkmdate,postingflg " & _ 
 " from bkmhdr " & _ 
 " where DOCUMENT_CLASSIFICATION <> '505' and gangcode='" & gangcode & "' and bkmdate=TO_DATE('" & tglbkm & "','MM/DD/YYYY') AND FCBA='" & fcba & "'"
 
end if

if employee = 101 then

 sqlview =  "select bkmno, sectioncode,gangcode,bkmdate,postingflg " & _ 
 " from bkmhdr " & _ 
 " where DOCUMENT_CLASSIFICATION = '505' and gangcode='" & gangcode & "' and bkmdate=TO_DATE('" & tglbkm & "','MM/DD/YYYY') AND FCBA='" & fcba & "'"
 
end if

'AMBIL EMPLOYE DARI INPUT RAWATAN 
if employee = 200 then

 sqlview =  "select bkmno, sectioncode,gangcode,bkmdate " & _ 
 " from bkmhdr " & _ 
 " where gangcode='" & gangcode & "' AND vehiclecode='" & vehiclecode & "' and bkmdate=TO_DATE('" & tglbkm & "','MM/DD/YYYY') AND FCBA='" & fcba & "'"
 
end if

'AMBIL EMPLOYE DARI INPUT RAWATAN
if employee = 1 then


sqlview = "SELECT " & _ 
"    e.fccode,  " & _ 
"    e.fcname || '-' || e.fccode fcname, " & _ 
"    e.presentgrade, " & _ 
"    e.fcba fcbaname,  " & _ 
"    TO_CHAR(timein,'dd/mm/yyyy hh24:mi:ss') timeinx, " & _ 
"    TO_CHAR(timeout,'dd/mm/yyyy hh24:mi:ss') timeoutx,  " & _ 
"    NVL(TO_CHAR(a.timein,'HH24:MI:SS'),'-') timein, " & _ 
"    NVL(TO_CHAR(a.timeout,'HH24:MI:SS'),'-') timeout,  " & _ 
"    bkm.fcba,  " & _ 
"    bkm.employeecode, " & _ 
"    NVL(bkm.ttlmandays,0) ttlmandays " & _ 
" FROM employee e  " & _ 
" LEFT JOIN USERFINGER.ATTENDANCE_TAB a " & _ 
"       ON a.EMPNO = e.FCCODE " & _ 
"      AND a.FCBA  = e.FCBA  " & _ 
"      AND a.TDATE = TO_DATE('" & tglbkm & "','MM/DD/YYYY') " & _ 
" LEFT JOIN ( " & _ 
"            SELECT  " & _ 
"                h.fcba, " & _ 
"                d.employeecode, " & _ 
"                NVL(SUM(d.mandays),0) ttlmandays " & _ 
"            FROM BKMUPAHDTL d  " & _ 
"            INNER JOIN bkmhdr h " & _ 
"                    ON h.BKMNO = d.BKMNO  " & _ 
"                   AND h.fcba  = d.fcba " & _ 
"            INNER JOIN ( " & _ 
"                SELECT * " & _ 
"                FROM ( " & _ 
"                    SELECT es.*, " & _ 
"                           ROW_NUMBER() OVER( " & _ 
"                               PARTITION BY es.fcba  " & _ 
"                               ORDER BY es.start_date DESC " & _ 
"                           ) rn " & _ 
"                    FROM EMPALLDEDSTATUS es " & _ 
"                    WHERE TO_DATE('" & tglbkm & "','MM/DD/YYYY')  " & _ 
"                          BETWEEN es.START_DATE AND es.END_DATE " & _ 
"                ) x " & _ 
"                WHERE x.rn = 1 " & _ 
"            ) es ON es.fcba = h.fcba " & _ 
"            WHERE h.fcba = '" & fcba & "' " & _ 
"              AND TRUNC(h.bkmdate,'MM') = TRUNC(TO_DATE('" & tglbkm & "','MM/DD/YYYY'),'MM') " & _ 
"            GROUP BY h.fcba, d.employeecode " & _ 
"          ) bkm  " & _ 
"       ON e.fcba = bkm.fcba  " & _ 
"      AND e.fccode = bkm.employeecode " & _ 
"WHERE e.gangcode = '" & gangcode & "' " & _ 
"  AND e.fcba = '" & fcba & "' " & _ 
"  AND e.dateterminate IS NULL " & _ 
"ORDER BY e.fcname ASC"

'response.write sqlview
end if


'AMBIL EMPLOYE DARI INPUT PANEN
if employee = 2 then

' sqlview =  "select e.fccode, e.fcname || '-' ||  e.fccode as fcname, e.fcba,to_char(timein,'dd/mm/yyyy hh24:mi:ss') timeinx,to_char(timeout,'dd/mm/yyyy hh24:mi:ss') timeoutx,NVL(to_char(a.timein,'HH24:MI:SS'),'-')timein,NVL(to_char(a.timeout,'HH24:MI:SS'),'-')timeout " & _
 '" from employee e left join USERFINGER.ATTENDANCE_TAB a ON a.EMPNO=e.FCCODE AND a.FCBA=e.FCBA AND a.TDATE=TO_DATE('" & tglbkm & "','MM/DD/YYYY') " & _
 '"  where e.gangcode='" & gangcode & "' and e.fcba='" & fcba & "' and e.dateterminate is null order by e.fcname asc"
 'response.write sqlview

 sqlview = "select e.fccode, e.fcname || '-' ||  e.fccode fcname,e.presentgrade,e.fcba fcbaname,   to_char(timein,'dd/mm/yyyy hh24:mi:ss') timeinx,to_char(timeout,'dd/mm/yyyy hh24:mi:ss') timeoutx, " & _ 
" NVL(to_char(a.timein,'HH24:MI:SS'),'-')timein,NVL(to_char(a.timeout,'HH24:MI:SS'),'-')timeout,  bkm.* from employee e " & _ 
" left join USERFINGER.ATTENDANCE_TAB a ON a.EMPNO=e.FCCODE AND a.FCBA=e.FCBA AND a.TDATE=TO_DATE('" & tglbkm & "','MM/DD/YYYY') " & _ 
" left join " & _ 
"  (select h.fcba,d.employeecode,sum(d.mandays) sumHK  " & _ 
"  from BKMUPAHDTL d " & _ 
"  inner join bkmhdr h on (h.BKMNO=d.BKMNO and h.fcba=d.fcba) " & _ 
"  inner join EMPALLDEDSTATUS es ON ( h.fcba=es.fcba AND  TO_DATE('" & tglbkm & "','MM/DD/YYYY') BETWEEN START_DATE	 and END_DATE )  " & _ 
"  where h.fcba='" & fcba & "'   " & _ 
"  /*AND es.month=to_char(to_date('" & tglbkm & "','MM/DD/YYYY'),'MM') and es.year=to_char(to_date('" & tglbkm & "','MM/DD/YYYY'),'YYYY') */ " & _ 
"  and bkmdate between es.START_DATE	 and es.END_DATE " & _ 
"  /*and bkmdate < TO_DATE('" & tglbkm & "','MM/DD/YYYY') */ " & _ 
"  group by  h.fcba,d.employeecode " & _ 
"   )bkm on (e.fcba=bkm.fcba and e.fccode=bkm.employeecode ) " & _ 
"   where gangcode='" & gangcode & "' and e.fcba='" & fcba & "'  " & _ 
"   and e.dateterminate is null " & _ 
" order by fcname asc"

end if


end if

'AMBIL HK EMPLOYE HI
if employee = 300 then

 sqlview =  "select nvl(sum(d.mandays),0)mandayshi from bkmhdr h inner join BKMUPAHDTL d on (h.bkmno=d.bkmno and h.fcba=d.fcba) where h.fcba='" & fcba & "' AND h.bkmdate=to_date('" & tglbkm & "','mm/dd/yyyy') and d.employeecode='" & employeecode & "' "
' response.write sqlview

end if

if employee = 301 then

 sqlview =  "select nvl(sum(d.mandays),0)mandayshi from bkmhdr h inner join BKMUPAHDTL d on (h.bkmno=d.bkmno and h.fcba=d.fcba) where h.POSTINGFLG='1' and  h.fcba='" & fcba & "' AND h.bkmdate=to_date('" & tglbkm & "','mm/dd/yyyy') and d.employeecode='" & employeecode & "' "
' response.write sqlview

end if

'ambil data dr master bankstock
if employee = 400 then

 sqlview =  "select fccode from bankstock where stocksubgroupcode='" & stocksubgroupcode & "' order by fccode desc FETCH FIRST 1 ROWS ONLY "
' response.write sqlview

end if

'ambil data dr lastupdate closingunit
if employee = 410 then


    if modulid = 1 then 
		modulname = "aresta_closing"
	end if
	if modulid = 2 then 
		modulname = "kasbank_closing"
	end if
	if modulid = 3 then 
		modulname = "stock_closing"
	end if
	if modulid = 4 then 
		modulname = "wip_closing"
	end if
	if modulid = 5 then 
		modulname = "wsvra_closing"
	end if
	if modulid = 6 then 
		modulname = "payroll_closing"
	end if
	if modulid = 7 then 
		modulname = "unit_closing"
	end if

 sqlview =  "select TO_CHAR(lastupdate,'DD/MM/YYYY HH24:MI:SS') as lastupdate from EMPALLDED_UNIT_CLOSING where FCBA='" & fcba & "' AND YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND lower(MODULENAME)='" & modulname & "' order by lastupdate FETCH FIRST 1 ROWS ONLY "
'' response.write sqlview

end if

'response.write sqlview
 
Set List = Server.CreateObject("ADODB.Recordset")
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