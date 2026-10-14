<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<link rel="stylesheet" href="css/jquery-ui.css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<title>Reporting - Monitoring PR</title>
</head>
<%
Dim v_statusgrn,v_tahun,v_company,v_fcba,v_CREATIONDATE_From,v_CREATIONDATE_To,v_document_classification,filter_doc,filtercompany1,filterStatusGRN,filter 
dim whereClause,doc_classification
v_statusgrn  = trim(REQUEST("statusgrn"))
v_company    = trim(REQUEST("pcompany"))
v_fcba       = trim(REQUEST("pfcba"))
v_CREATIONDATE_From       = trim(REQUEST("date_From"))
v_CREATIONDATE_To       = trim(REQUEST("date_To"))
v_document_classification       = trim(REQUEST("document_classification"))

%>
<script src="/distexcel/tableToExcel.js"></script>
<script type="text/javascript">

	function getExport()
	{
		var company 	= '<%= v_company %>';
		var nm = 'MonitoringPR_'+ 'PT'+ company +".xlsx"; 
	//alert("nm");
	 
		TableToExcel.convert(document.getElementById("product-table"), {
            name: nm,
            sheet: {
            name: "Monitoring_PR"
            }
          });
		//$('table').tblToExcel();
		
	}
</script>
<script type="text/javascript" >
 $(function() {	
  setTimeout(function () {
        getExport();
     }, 1000); 

});

</script>

<%
userid = session("CodeName")
IF UCase(userid) = "DIV_HEAD_COMM" or UCase(userid) ="DEPT_HEAD_PURCH" or UCase(userid) ="STAFF_OFFICER_PURCH_02" or UCase(userid) ="STAFF_OFFICER_PURCH" or UCase(userid) ="SISKA_MEILINA_SEMBIRING" or UCase(userid) ="RINTO_PURNAMASIDI" THEN
	
	whereClause = " WHERE b.prdate BETWEEN TO_DATE('" & v_CREATIONDATE_From & "','MM/DD/YYYY') AND TO_DATE('" & v_CREATIONDATE_To & "','MM/DD/YYYY') AND b.rowstateid not in ('Rejected') AND b.document_classification not in('206','210') "
ELSE 
	whereClause = " WHERE b.prdate BETWEEN TO_DATE('" & v_CREATIONDATE_From & "','MM/DD/YYYY') AND TO_DATE('" & v_CREATIONDATE_To & "','MM/DD/YYYY') AND b.rowstateid not in ('Rejected') "
	
END IF
'whereClause = " WHERE b.prdate BETWEEN TO_DATE('" & v_CREATIONDATE_From & "','MM/DD/YYYY') AND TO_DATE('" & v_CREATIONDATE_To & "','MM/DD/YYYY') AND b.rowstateid not in ('Rejected') "
 	if trim(v_company)="all" THEN
		filtercompany1 = " and a.fcba in(select fccode from businessunit b  where  b.plantation ='FBG') "
	else
		if trim(v_fcba) = "all" then 
			filtercompany1 = " and a.fcba in ( select fccode from businessunit where fccompanycode='" & v_company & "' )"
		else 
			filtercompany1 = " and  a.fcba ='" & v_fcba & "' "
		end if
    end if
	If Trim(v_statusgrn) = "donegrn" Then
		filterStatusGRN = " AND e.receivenotecode IS NOT NULL " ' Sudah GRN
			Else
				If Trim(v_statusgrn) = "notgrn" Then
					filterStatusGRN = " AND e.receivenotecode IS NULL " ' Belum GRN
			Else
					filterStatusGRN = ""
	End If

	if v_document_classification<>"all" then 
		filter_doc = " AND b.DOCUMENT_CLASSIFICATION= '" & v_document_classification & "'"
	end if
end if
		If filtercompany1 <> "" Then whereClause = whereClause & filtercompany1
		If filterStatusGRN <> "" Then whereClause = whereClause & filterStatusGRN	
		

sqlview = " WITH PRD_HIST AS ( " & _
" SELECT prcode, fcba, lineid, fcentry, lastupdate AS lupdate, currentauthorizelevel, " & _
" ROW_NUMBER() OVER (PARTITION BY prcode, fcba, lineid, currentauthorizelevel ORDER BY lastupdate DESC) AS rn " & _
" FROM purchaserequestdetails_hist ) " & _
" SELECT DISTINCT a.prcode " & _
"      ,    b.prdate " & _
"      ,	b.rowstateid " & _
"      ,	a.itemcode " & _
"      ,	x.fcname AS stockname " & _
"      ,    x.UNITOFMEASUREMENT uom " & _
"      ,    a.requestedqty " & _
"      ,    a.approvedqty " & _  
"      ,    sp.fcname suppliercode  " & _    
"      ,    d.pocode  " & _
"      ,    d.podate  " & _
"      ,    c.quantity " & _
"      ,    nvl(c.unitprice,0)unitprice " & _
"      ,    e.receivenotecode  " & _
"      ,    e.rndate, TO_CHAR(e.LASTUPDATE, 'mm/dd/yyyy') AS GRNLASTUPDATE  " & _
"      ,    f.quantity qtyGRN " & _
"	   ,    dc.fcname" &  _
"      ,    a.description "  & _
"      ,	NVL(c.keterangan, a.keterangan) as keterangan  " & _
"      ,    nvl(c.keterangan1, a.keterangan1) as keterangan1 "  & _
"      ,    a.lineid "  & _
"      ,    c.linenokey linenokeypod "  & _
"      ,    a.fcba "  & _
"      ,    bu.FCCOMPANYCODE "  & _
"      ,    h.fcentry AS fcentry_161 "  & _
"	   ,	h.lupdate AS approve_date_161 "  & _
"	   ,	i.fcentry AS fcentry_190 "  & _
"	   ,	i.lupdate AS approve_date_190 "  & _
"	   ,	j.fcentry AS fcentry_900 "  & _
"	   ,	j.lupdate AS approve_date_900 "  & _
"	   ,	bu.region AS region "  & _
"	   ,	bu.fccompanycode AS fccompanycode "  & _
"			FROM bankstock x "  & _
"			INNER JOIN purchaserequestdetails a ON x.fccode = a.itemcode "  & _
"			INNER JOIN purchaserequest b ON a.prcode = b.prcode AND a.fcba = b.fcba "  & _
"			LEFT JOIN PRD_HIST h ON b.prcode = h.prcode AND a.fcba = h.fcba AND a.lineid = h.lineid "  & _
"		    AND h.currentauthorizelevel = '161' AND h.rn = 1"  & _
"			LEFT JOIN PRD_HIST i ON a.prcode = i.prcode AND a.fcba = i.fcba AND a.lineid = i.lineid "  & _
"		    AND i.currentauthorizelevel = '190' AND i.rn = 1"  & _
"			LEFT JOIN PRD_HIST j ON b.prcode = j.prcode AND a.fcba = j.fcba AND a.lineid = j.lineid "  & _
"    		AND j.currentauthorizelevel = '900' AND j.rn = 1"  & _
"			LEFT JOIN purchaseorderdetails c ON a.prcode = c.prcode AND a.lineid = c.prlineid AND a.itemcode = c.itemcode "  & _
"			LEFT JOIN purchaseorder d ON c.pocode = d.pocode AND c.fcba = d.fcba "  & _
"         	LEFT JOIN supplier sp" & _
"            ON (d.suppliercode = sp.fccode AND d.fcba = sp.fcba)" & _
"			LEFT JOIN receivenotedetail f ON c.linenokey = f.purchaselinenokey AND a.itemcode = f.purchaseitemcode "  & _
"			LEFT JOIN receivenote e ON f.receivenotecode = e.receivenotecode AND f.fcba = e.fcba "  & _
"		    AND e.fcba = f.fcba "  & _
"		    AND c.linenokey = f.purchaselinenokey "  & _
"		    AND a.itemcode = f.purchaseitemcode"  & _
"			LEFT JOIN ( "  & _
"				SELECT DISTINCT fccode, fcname "  & _
"				FROM document_classification WHERE DOCUMENT_TYPE='PR' "  & _
"			) dc ON dc.fccode = b.document_classification "  & _
"			LEFT JOIN BUSINESSUNIT BU ON BU.FCCODE=b.fcba" & _
				whereClause & _
				filter_doc & _
"           ORDER BY bu.region,bu.fccompanycode, a.prcode,b.prdate,a.itemcode  "
set rsview = connStr.Execute(sqlview)
' response.Write(sqlview)
response.Flush()
Set rsCat = Server.CreateObject("ADODB.Recordset")
			rsCat.CursorType = 3
			rsCat.CursorLocation = 3
			rsCat.Open sqlview, connstr

			dim sqluser,rsuser,menu_po
			sqluser = "select menu_po from users_tab where fccode='" & session("CodeName") & "'"
								set rsuser = connStr.Execute(sqluser)
								if not rsuser.eof then
									menu_po = rsuser("menu_po")
								end if
			%>

<body>
    <span id="loadingjudul1">Proses export data...</span>
    <div id="tableview" style="display:none;">
   	<table border="0" width="100%" id="product-table" >
			<thead>
				<tr>
					<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">MONITORING PR <%=v_company%> PT : <%=v_company%> STATUS GRN : <%=v_statusgrn%> PERIODE  : <%=v_CREATIONDATE_From%> - <%=v_CREATIONDATE_To%></th>		
				</tr>
				<tr >
							<th rowspan="2" >COMPANY</th>
							<th rowspan="2" >FCBA</th>
                            <%
                                dim userlogins11
                                userlogins11 = Session("CodeName")
                                if userlogins11="STAFF_OFFICER_PURCH" OR userlogins11="STAFF_OFFICER_PURCH_02" OR userlogins11="DIV_HEAD_COMM" OR userlogins11="IT_HO6" OR userlogins11="IT_HO2" OR userlogins11="SISKA_MEILINA_SEMBIRING" OR userlogins11="RINTO_PURNAMASIDI" OR userlogins="DEPT_HEAD_PURCH" THEN
                            %>
                                <th colspan="15" >PR</th>
                            <%else%> 
                                <th colspan="15" >PR</th>
                            <% end if%>
                            <%
                                dim userlogins111
                                userlogins111 = Session("CodeName")
                                if userlogins111="STAFF_OFFICER_PURCH" OR userlogins111="STAFF_OFFICER_PURCH_02" OR userlogins111="DIV_HEAD_COMM" OR userlogins111="IT_HO6" OR userlogins="IT_HO2" OR userlogins="SISKA_MEILINA_SEMBIRING" OR userlogins111="RINTO_PURNAMASIDI" OR userlogins111="DEPT_HEAD_PURCH" THEN
                            %>
                                <th colspan="6" >PO</th>
                            <%else%> 
                                <th colspan="6" >PO</th>
                            <% end if%>
                            <%
                                dim userlogin112
                                userlogin112 = Session("CodeName")
                                if userlogin112="STAFF_OFFICER_PURCH" OR userlogin112="STAFF_OFFICER_PURCH_02" OR userlogin112="DIV_HEAD_COMM" OR userlogin112="IT_HO6" OR userlogin112="IT_HO2" OR userlogin112="SISKA_MEILINA_SEMBIRING" OR userlogin112="RINTO_PURNAMASIDI" OR userlogin112="DEPT_HEAD_PURCH" THEN
                            %>
                                <th colspan="6" >GRN</th>
                            <%else%> 
                                -
                            <% end if%>
							<th rowspan="2" >KETERANGAN</th>
							<th rowspan="2" >HISTORI KETERANGAN</th>
                </tr>
                <tr>
                    <th style="color:#fff;">NO.PR</th>
                    <th style="color:#fff;">TGL.PP</th>
					<th style="color:#fff;">PR Line</th>
                    <th style="color:#fff;">Kode.Barang</th>
                    <th style="color:#fff;">Nama.Barang</th>
                    <th style="color:#fff;">Satuan</th>
					<th style="color:#fff;">Keterangan</th>
					<th style="color:#fff;">Qty Diminta</th>
                    <th style="color:#fff;">Qty Approve</th>
					<th style="color:#fff;">Classification</th>
					<th style="color:#fff;">Status</th>
                    <th style="color:#fff;">Approve 1</th>
					<th style="color:#fff;">Approve 2</th>
					<th style="color:#fff;">Approve 3</th>
                    <th style="color:#fff;">Dura</th>
                    <th style="color:#fff;">Supplier</th>
                    <th style="color:#fff;">NO.PO</th>
                    <th style="color:#fff;">Tgl.PO</th>
                    <%
                        dim podura
                        podura = Session("CodeName")
                        if podura="STAFF_OFFICER_PURCH" OR podura="STAFF_OFFICER_PURCH_02" OR podura="DIV_HEAD_COMM" OR podura="IT_HO6" OR podura="IT_HO2" OR podura="SISKA_MEILINA_SEMBIRING" OR podura="RINTO_PURNAMASIDI" OR podura="DEPT_HEAD_PURCH" THEN
                    %>
                        <th style="color:#fff;">Dura</th>
                    <%else%> 
                        -
                    <% end if%>
                        <th style="color:#fff;">Qty PO</th>
                    <%if menu_po="TRUE" then%>
                        <th >Price PO</th>
					<%END IF%>
                    
                    <th style="color:#fff;">NO.GRN</th>
                    <th style="color:#fff;">Tgl.GRN</th>

                    <%
                        dim userlogin122
                        userlogin122 = Session("CodeName")
                        if userlogin122="STAFF_OFFICER_PURCH" OR userlogin122="STAFF_OFFICER_PURCH_02" OR userlogin122="DIV_HEAD_COMM" OR userlogin122="IT_HO6" OR userlogin122="IT_HO2" OR userlogin122="SISKA_MEILINA_SEMBIRING" OR userlogin122="RINTO_PURNAMASIDI" OR userlogin122="DEPT_HEAD_PURCH" THEN
                    %>
                        <th style="color:#fff;">Dura</th>
                        <th style="color:#fff;">LU.GRN</th>
					    <th style="color:#fff;">Dura</th>
                    <%else%> 
                        -
                    <% end if%>
                    <th style="color:#fff;">Qty GRN</th>
                </tr>
		</thead>
<tbody>
<%
    dim no, userapproved1, userapproved2, userapproved3, tglapp1, tglapp2, tglapp3, statusauto
    no = 1
    If Not rsview.EOF Then
        While Not rsview.EOF
            statusauto = rsview("rowstateid")
            userapproved1 = rsview("fcentry_161") 
            userapproved2 = rsview("fcentry_190")
            userapproved3 = rsview("fcentry_900")
            tglapp1 = rsview("approve_date_161") 
            tglapp2 = rsview("approve_date_190")
            tglapp3 = rsview("approve_date_900")
            If Trim(rsview("prcode")) <> "" Then 
%>
        <tr>
            <td align="center"><font face="Verdana" size="0.5"><%=rsview("fccompanycode")%></td>
            <td align="center"><font face="Verdana" size="0.5"><%=rsview("fcba")%></td>
            <td align="center"><font face="Verdana" size="0.5"><%=rsview("prcode")%></td>
            <td align="center"><font face="Verdana" size="0.5"><%=rsview("prdate")%></td>
            <td align="center"><font face="Verdana" size="0.5"><%=rsview("lineid")%></td>
            <td align="center"><font face="Verdana" size="0.5"><%=rsview("itemcode")%></td>
            <td align="center"><font face="Verdana" size="1"><%=rsview("stockname")%></td>
            <td align="center"><font face="Verdana" size="1"><%=rsview("uom")%></td>
            <td align="center"><font face="Verdana" size="1"><%=rsview("description")%></td>
            <td align="center"><font face="Verdana" size="1"><%=rsview("requestedqty")%></td>
            <td align="right"><font face="Verdana" size="1"><%=rsview("approvedqty")%></td>
            <td align="center"><font face="Verdana" size="1"><%=rsview("fcname")%></td>
            <td align="center"><font face="Verdana" size="1"><%=statusauto%></td>
            <td align="center"><font face="Verdana" size="1"><%=userapproved1 & " - " & tglapp1%></td>
            <td align="center"><font face="Verdana" size="1"><%=userapproved2 & " - " & tglapp2%></td>
            <td align="center"><font face="Verdana" size="1"><%=userapproved3 & " - " & tglapp3%></td>
            <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1">
                <%
                    dim prdura
                    prdura = Session("CodeName")
                    if prdura="STAFF_OFFICER_PURCH" OR prdura="STAFF_OFFICER_PURCH_02" OR prdura="DIV_HEAD_COMM" OR prdura="IT_HO2" OR prdura="SISKA_MEILINA_SEMBIRING" OR prdura="RINTO_PURNAMASIDI" OR prdura="DEPT_HEAD_PURCH" THEN
                        if Not IsNull(tglapp3) and tglapp3 <> "" then
                            Response.Write DateDiff("d", rsview("prdate"), tglapp3)
                        else
                            Response.Write "-"
                        end if
                    else 
                        Response.Write "-"
                    end if
                %>
            </td>
            <td align="center"   style="<%=cmdstyle%>"><font face="Verdana" size="1">
                <%
                    dim userlogin13
                    userlogin13 = Session("CodeName")
                    if userlogin13="STAFF_OFFICER_PURCH" OR userlogin13="STAFF_OFFICER_PURCH_02" OR userlogin13="DIV_HEAD_COMM" OR userlogin13="IT_HO6" OR userlogin13="SISKA_MEILINA_SEMBIRING" OR userlogin13="RINTO_PURNAMASIDI" OR userlogin13="DEPT_HEAD_PURCH" THEN
                %>
                <%=rsview("suppliercode")%>
                <%else%> 
                -
                <% end if%>
			</td>
            <td align="center"><font face="Verdana" size="1"><%=rsview("pocode")%></td>
            <td align="center"><font face="Verdana" size="1"><%=rsview("podate")%></td>
            <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1">
                <%
                    dim podura1
                    podura1 = Session("CodeName")
                    if podura1="STAFF_OFFICER_PURCH" OR podura1="STAFF_OFFICER_PURCH_02" OR podura1="DIV_HEAD_COMM" OR podura1="IT_HO6" OR podura1="IT_HO2" OR podura1="SISKA_MEILINA_SEMBIRING" OR podura1="RINTO_PURNAMASIDI" OR podura1="DEPT_HEAD_PURCH" THEN
                %>
                    <%=DateDiff("d", tglapp3, rsview("podate"))%>
                <%else%> 
                    -
                <% end if%>
            </td>
            <td align="right"><font face="Verdana" size="1"><%=rsview("quantity")%></td>
            <%if menu_po="TRUE" then%>
            <td align="Right"  style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rsview("unitprice"),0)%></td>
            <%end if%>
            <td align="center"><font face="Verdana" size="1"><%=rsview("receivenotecode")%></td>
            <td align="center"><font face="Verdana" size="1"><%=rsview("rndate")%></td>
            <%
                dim userlogin1
                userlogin1 = Session("CodeName")
                if userlogin1="STAFF_OFFICER_PURCH" OR userlogin1="STAFF_OFFICER_PURCH_02" OR userlogin1="DIV_HEAD_COMM" OR userlogin1="IT_HO6" OR userlogin1="IT_HO2" OR userlogin1="SISKA_MEILINA_SEMBIRING" OR userlogin1="RINTO_PURNAMASIDI" OR userlogin1="DEPT_HEAD_PURCH" THEN
            %>
                <td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=DateDiff("d", rsview("podate"), rsview("rndate"))%> </td>
                <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("GRNLASTUPDATE")%> </td>
                <td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=DateDiff("d", rsview("rndate"), rsview("GRNLASTUPDATE"))%> </td>
            <%else%> 
                -
            <% end if%>
            <td align="Right"  style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("qtyGRN")%></td>
            <td align="left"><font face="Verdana" size="1"><%=rsview("keterangan")%></td>
            <td align="left"><font face="Verdana" size="1"><%=rsview("keterangan1")%></td>
        </tr>
<%
            End If 
            no = no + 1
            rsview.MoveNext
        Wend
    End If
    rsview.Close()
    rsview = Nothing
%>
</tbody>
</table>
</div>

</BODY>
</HTML>
