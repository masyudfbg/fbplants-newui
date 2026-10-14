
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


dim txtcari,fcba,txtsupplier ,bulansplit,tglsplit,thnsplit,blnthn,fcbaselect,section
txtcari = Request("keyword")
fcba = session("fcba")
jmlitem = Request("jmlitem")
tahun = Request("tahun")
bulan = Request("bulan")
contractorcode = Request("contractorcode")
fcbaselect = Request("fcbaselect")
activitywip = Request("activitywip")
agreement = Request("agreement")
linenowip = Request("linenowip")
bkmno = Request("bkmno")
tglbkm = Trim(Request("tglbkm"))
arrTanggal = Split(tglbkm, "/")

If UBound(arrTanggal) = 2 Then
    bulansplit = CInt(arrTanggal(0))
    tglsplit   = CInt(arrTanggal(1))
    thnsplit   = CInt(arrTanggal(2))
End If

blnthn      = thnsplit
section     = Request("section")

Set List = Server.CreateObject("ADODB.Recordset")

if jmlitem = 1 then 

sqlview =  "SELECT distinct d.ITEMCODE, d.ITEMCODE || ' - ' || s.fcname " & _
"FROM BKMHDR h,BKMBAHANDTL d,stock s " & _
"WHERE h.bkmno=d.bkmno and h.fcba=d.fcba and s.fccode=d.ITEMCODE and s.fcba=d.fcba and h.bkmno='" & bkmno & "'"
end if 


if jmlitem = 0 then

    sqlview = "SELECT DISTINCT" & _
                "    x.FCCODE as id, x.FCCODE || ' - ' || x.fcname || ' - STOK : ' || x.stock_qty AS text " & _
                "FROM ( " & _
                "    SELECT " & _
                "        stck.FCCODE, " & _
                "        stck.fcname, " & _
                "        sl.OPENINGBALANCEQTY + (NVL(rn_sum.total_in, 0) + NVL(rcv_sum.total_in, 0) + NVL(trsd_sum.total_inrcv, 0)) AS stock_qty " & _
                "    FROM stockledger sl " & _
                "    LEFT JOIN ( " & _
                "        SELECT FCCODE, MIN(fcname) AS fcname " & _
                "        FROM STOCK " & _
                "        GROUP BY FCCODE " & _
                "    ) stck ON sl.ITEMCODE = stck.FCCODE " & _
                "    LEFT JOIN ( " & _
                "        SELECT rnd.PURCHASEITEMCODE AS itemcode, SUM(rnd.QUANTITY) AS total_in " & _
                "        FROM receivenotedetail rnd " & _
                "        JOIN receivenote rn ON rnd.RECEIVENOTECODE = rn.RECEIVENOTECODE " & _
                "        WHERE rn.RNDATE BETWEEN TO_DATE('"&trim(bulansplit)&"/01/"&trim(thnsplit)&"','mm/dd/yyyy') AND TO_DATE('"&trim(bulansplit)&"/"&trim(tglsplit)&"/"&trim(thnsplit)&"','mm/dd/yyyy') AND rn.FCBA = '"&fcba&"' " & _
                "        AND rnd.LOCATIONCODE IN (SELECT DISTINCT FCCODE FROM store WHERE SECTIONNAME = '"&section&"' and fcba = '"& fcba &"') " &_
                "        GROUP BY rnd.PURCHASEITEMCODE " & _
                "    ) rn_sum ON sl.ITEMCODE = rn_sum.itemcode " & _
                "    LEFT JOIN ( " & _
                "        SELECT rcvd.STOCKCODE AS itemcode, SUM(rcvd.QUANTITYRECEIVED) AS total_in " & _
                "        FROM RECEIVEDTASKDTL_TAB rcvd " & _
                "        JOIN RECEIVEDTASK_TAB rcv ON rcvd.ORDERNO = rcv.ORDERNO " & _
                "        WHERE rcv.ORDERDATE BETWEEN TO_DATE('"&trim(bulansplit)&"/01/"&trim(thnsplit)&"','mm/dd/yyyy') AND TO_DATE('"&trim(bulansplit)&"/"&trim(tglsplit)&"/"&trim(thnsplit)&"','mm/dd/yyyy') AND rcv.FCBA = '"&fcba&"' " & _
                "        AND rcvd.LOCATIONCODE IN (SELECT DISTINCT FCCODE FROM store WHERE SECTIONNAME = '"&section&"' and fcba = '"& fcba &"') " &_
                "        GROUP BY rcvd.STOCKCODE " & _
                "    ) rcv_sum ON sl.ITEMCODE = rcv_sum.itemcode " & _
                "    LEFT JOIN " & _
                "       (SELECT rcvdivdtl.STOCKCODE AS itemcode, " & _
                "               SUM(rcvdivdtl.QUANTITYTRANSFERED) AS total_out " & _
                "       FROM TRANSPORTTASKDIVDTL_TAB rcvdivdtl " & _
                "       JOIN TRANSPORTTASKDIV_TAB rcvdiv ON rcvdivdtl.ORDERNO = rcvdiv.ORDERNO " & _
                "       WHERE rcvdiv.ORDERDATE BETWEEN TO_DATE('"&trim(bulansplit)&"/01/"&trim(thnsplit)&"','mm/dd/yyyy') AND TO_DATE('"&trim(bulansplit)&"/"&trim(tglsplit)&"/"&trim(thnsplit)&"','mm/dd/yyyy') AND rcvdiv.FCBA = '"&fcba&"' " & _
                "           AND rcvdivdtl.STORECODEFROM = 'ST1'  AND rcvdivdtl.STORECODETO IN (SELECT FCCODE FROM store WHERE SECTIONNAME = '"&section&"' and fcba = '"& fcba &"') " & _
                "       GROUP BY rcvdivdtl.STOCKCODE) rcvdivdtlsum ON sl.ITEMCODE = rcvdivdtlsum.itemcode " & _
                "    LEFT JOIN ( " & _
                "        SELECT isd.STOCKCODE AS itemcode, SUM(isd.QUANTITY) AS total_out " & _
                "        FROM ISSUEDETAILS isd " & _
                "        JOIN ISSUE issue ON isd.SIVCODE = issue.SIVCODE " & _
                "        WHERE issue.SIVDATE BETWEEN TO_DATE('"&trim(bulansplit)&"/01/"&trim(thnsplit)&"','mm/dd/yyyy') AND TO_DATE('"&trim(bulansplit)&"/"&trim(tglsplit)&"/"&trim(thnsplit)&"','mm/dd/yyyy') AND isd.FCBA = '"&fcba&"' " & _
                "        AND issue.STORECODE IN (SELECT DISTINCT FCCODE FROM store WHERE SECTIONNAME = '"&section&"' and fcba = '"& fcba &"')" & _
                "        GROUP BY isd.STOCKCODE " & _
                "    ) isv_sum ON sl.ITEMCODE = isv_sum.itemcode " & _
                "    LEFT JOIN ( " & _
                "        SELECT trsd.STOCKCODE AS itemcode, SUM(trsd.QUANTITYTRANSFERED) AS total_inrcv " & _
                "        FROM TRANSPORTTASKDIVDTL_TAB trsd " & _
                "        JOIN TRANSPORTTASKDIV_TAB trs ON trsd.ORDERNO = trs.ORDERNO " & _
                "        WHERE trs.ORDERDATE BETWEEN TO_DATE('"&trim(bulansplit)&"/01/"&trim(thnsplit)&"','mm/dd/yyyy') AND TO_DATE('"&trim(bulansplit)&"/"&trim(tglsplit)&"/"&trim(thnsplit)&"','mm/dd/yyyy') AND trs.FCBA = '"&fcba&"' " & _
                "        AND trsd.STORECODETO IN (SELECT DISTINCT FCCODE FROM store WHERE SECTIONNAME = '"&section&"' and fcba = '"& fcba &"') " &_
                "        GROUP BY trsd.STOCKCODE " & _
                "    ) trsd_sum ON sl.ITEMCODE = trsd_sum.itemcode " & _
                "    WHERE sl.storecode IN (SELECT FCCODE FROM store WHERE SECTIONNAME = '"&section&"' and fcba = '"& fcba &"') " & _
                "      AND sl.FCBA = '"&fcba&"' " & _
                "      AND sl.YEAR = "&trim(thnsplit)&" " & _
                "      AND sl.MONTH = "&trim(bulansplit)&" " & _
                ") x " & _
                "WHERE x.fcname IS NOT NULL AND x.stock_qty > 0"

            ' Response.write sqlview
end if


' GET DATA MASTER VEHICLE'
if jmlitem = 2 then

sqlview =  "Select fccode, v.FCCODE as id, v.FCCODE || ' - ' || v.fcname as text " & _
" From   VEHICLE v  " & _
" WHERE v.FCCODE || ' - ' || v.fcname like UPPER('%" & txtcari & "%') AND v.FCBA='" & fcba & "' AND v.activation='Y' " & _ 
" /*UNION ALL */" & _ 
" /*select fccode, v.FCCODE as id, v.FCCODE || ' - ' || v.fcname as text */" & _
" /*from WORKSHOP v " & _
" /*where v.FCCODE || ' - ' || v.fcname like UPPER('%" & txtcari & "%') AND v.fcBA='" & fcba & "' */" & _
" ORDER BY FCCODE "
'Response.write sqlview
end if

' GET COUNT BKM VEHICLE'
if jmlitem = 3 then

sqlview =  "Select v.FCCODE as id, v.FCCODE || ' - ' || v.fcname as text " & _
" From   VEHICLE v INNER JOIN BKMVEHICLEDTL d ON(v.FCCODE=d.VEHICLECODE and v.FCBA=d.FCBA)  " & _
" WHERE d.bkmno='" & bkmno & "' ORDER BY v.FCCODE,v.FCNAME "

end if

' GET SHOW KATEGORI TRANSPORT
if jmlitem = 4 then

    sqlview =  "Select kt.FCCODE as id, kt.FCCODE || ' - ' || kt.fcname as text " & _
    " From   VEHICLEACTIVITYCATEGORY kt  " & _
    " WHERE kt.FCCODE || ' - ' || kt.fcname like ('%" & txtcari & "%') AND kt.fcba='" & fcba & "' AND kt.ISACTIVE=1 ORDER BY kt.FCNAME "

'response.write sqlview
end if 

' GET SHOW WIP
if jmlitem = 5 then

sqlview =  "Select wi.AGREEMENTCODE as id, wi.AGREEMENTCODE as text " & _
" From   CONTRACTAGREEMENT wi  " & _
" WHERE wi.FCBA='" & fcba & "' AND wi.CONTRACTORCODE='" & contractorcode & "' " & _ 
" AND TO_DATE('" & tglbkm & "','mm/dd/yyyy') Between  startdate AND finishdate " & _ 
" AND wi.AGREEMENTCODE like UPPER('%" & txtcari & "%') ORDER BY wi.AGREEMENTCODE "
'response.write sqlview
end if

' GET COUNT BKM WIP'
if jmlitem = 6 then

sqlview =  "Select wi.AGREEMENTCODE " & _
" From   CONTRACTAGREEMENT wi INNER JOIN BKMWIPDTL d ON(wi.AGREEMENTCODE=d.AGREEMENTCODE and wi.FCBA=d.FCBA)  " & _
" WHERE d.bkmno='" & bkmno & "' ORDER BY wi.AGREEMENTCODE"

end if

' GET SHOW CONTRACTOR
if jmlitem = 7 then

sqlview =  "Select ct.FCCODE as id, ct.FCCODE || ' - ' || ct.FCNAME as text " & _
" From   CONTRACTOR ct  " & _
" WHERE ct.FCBA='" & fcba & "' and ct.FCCODE || ' - ' || ct.fcname like UPPER('%" & txtcari & "%') ORDER BY ct.FCNAME "

end if

' GET SHOW LINENO IN CONTRACAGREEMENT
if jmlitem = 8 then

sqlview =  "SELECT LINENO id,LINENO text FROM CONTRACTAGREEMENTDETAIL " & _
" WHERE FCBA='" & fcba & "' AND JOBCODE='" & activitywip & "' AND AGREEMENTCODE='" & agreement & "' order by lineno asc "

end if

' GET SHOW RATE WIP IN CONTRACTAGREEMENT
if jmlitem = 9 then

sqlview =  "SELECT RATE FROM CONTRACTAGREEMENTDETAIL " & _
" WHERE FCBA='" & fcba & "' AND JOBCODE='" & activitywip & "' AND AGREEMENTCODE='" & agreement & "' " & _ 
" AND lineno='" & linenowip & "' "

end if

' GET SHOW CONTRACTOR
if jmlitem = 10 then

sqlview =  "select 'all' id,'--All--' text FROM Dual UNION Select ct.FCCODE as id, ct.FCCODE || ' - ' || ct.FCNAME as text " & _
" From   CONTRACTOR ct  " & _
" WHERE ct.FCBA='" & fcbaselect & "' and ct.FCCODE || ' - ' || ct.fcname like UPPER('%" & txtcari & "%') ORDER BY text "
'response.write sqlview
end if

' GET SHOW AGREEMENT
if jmlitem = 11 then

sqlview =  "select 'all' id,'--All--' text FROM Dual UNION Select ct.AGREEMENTCODE as id, ct.AGREEMENTCODE as text " & _
" From   CONTRACTAGREEMENT ct  " & _
" WHERE ct.FCBA='" & fcbaselect & "' and ct.CONTRACTORCODE='" & contractorcode & "' and ct.AGREEMENTCODE like UPPER('%" & txtcari & "%') ORDER BY text "

end if

' GET SHOW AGREEMENT RKB WIP
if jmlitem = 12 then

sqlview =  "select 'all' id,'--All--' text FROM Dual UNION Select ct.AGREEMENTCODE as id, ct.AGREEMENTCODE as text " & _
" From   CONTRACTAGREEMENT ct  " & _
" WHERE ct.FCBA='" & fcbaselect & "' and ct.AGREEMENTCODE like UPPER('%" & txtcari & "%') ORDER BY text "

end if

' GET SHOW AGREEMENT ALL
if jmlitem = 13 then

sqlview =  "select 'all' id,'--All--' text FROM Dual UNION Select ct.AGREEMENTCODE as id, ct.AGREEMENTCODE as text " & _
" From   CONTRACTAGREEMENT ct  " & _
" WHERE ct.FCBA='" & fcbaselect & "' and ct.AGREEMENTCODE like UPPER('%" & txtcari & "%') and EXTRACT(MONTH FROM ct.agreementdate) = "&bulan&" AND EXTRACT(YEAR FROM ct.agreementdate) = "&tahun&" ORDER BY text "
' response.write sqlview
end if


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