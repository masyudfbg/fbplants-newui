<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Purchase Request Approval</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>

</head>
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtPRDATE_From=trim(request("PRDATE_From"))
v_txtPRDATE_To=trim(request("PRDATE_To"))
v_txtPRCODE=trim(request("PRCODE"))
v_txtITEMCODE=trim(request("ITEMCODE"))
v_txtRequestorID=trim(request("REQUESTORID"))
v_txtRowStateId=trim(request("ROWSTATEID"))
businessunit=trim(request("businessunit"))

		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
			' response.write q
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
			end if

txtSQL = 	"  SELECT hdr.fcba," & _
			"         hdr.prcode," & _
			"         hdr.prdate," & _
			"         hdr.managedid," & _
			"         hdr.description," & _
			"         hdr.remarks," & _
			"         hdr.RequestorId," & _
			"         hdr.rowstateid," & _
			"         prd.ItemCode," & _
			"         st.fcname," & _
			"         prd.requestedqty," & _
			"         prd.approvedqty," & _
			"         prd.orderedqty," & _
			"         aus.authorizelevelcode," & _
			"         hdr.document_classification," & _
			"         NVL (prh.currentauthorizelevel, 0) currentauthorizelevel," & _
			"         aus.fcusercode NextApv" & _
			"    FROM purchaserequest hdr" & _
			" 		INNER JOIN purchaserequestdetails prd ON (hdr.prcode=prd.prcode AND hdr.fcba=prd.fcba)" & _
			"		INNER JOIN stock st ON (prd.itemcode = st.fccode AND prd.fcba=st.fcba)" & _ 
			"         LEFT JOIN authorizetemplate at" & _
			"            ON (    at.document_type = 'PR'" & _
			"                AND at.document_classification = hdr.document_classification)" & _
			"         LEFT JOIN" & _
			"         (  SELECT fcba," & _
			"                   prcode," & _
			"                   NVL (MAX (currentauthorizelevel), 0) currentauthorizelevel" & _
			"              FROM purchaserequestdetails_hist" & _
			"          GROUP BY fcba, prcode) prh" & _
			"            ON (prh.fcba = hdr.fcba AND prh.prcode = hdr.prcode)" & _
			"         INNER JOIN authorizestructure aus" & _
			"            ON (    aus.fcba = hdr.fcba" & _
			"                AND aus.authorizetemplatecode = at.fccode" & _
			"           /*     AND aus.fcusercode = '"& session("codeName") &"'*/" & _
			"                AND authorizestructure_api.get_nextlevel (" & _
			"                       'PR'," & _
			"                       hdr.document_classification," & _
			"                       hdr.fcba," & _
			"                       NVL (prh.currentauthorizelevel, 0)) =" & _
			"                       aus.authorizelevelcode)" & _
			"   WHERE hdr.rowstateid = 'AuthorizeOnProgress' and hdr.fcba = '"& businessunit &"'"

if  len(trim(v_txtPRDATE_From)) > 0 then
    txtSQL = txtSQL & " and hdr.prdate >=to_Date('" & v_txtPRDATE_From & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " From " & FormatDateTime(v_txtPRDATE_From,1)	
end if
if  len(trim(v_txtPRDATE_To)) > 0 then
    txtSQL =txtSQL &  " and hdr.prdate <=to_Date('" & v_txtPRDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " To " & FormatDateTime(v_txtPRDATE_To,1)	
end if
if  len(trim(v_txtPRCODE)) > 0 then
    txtSQL = txtSQL & " and hdr.prcode='" & v_txtPRCODE & "'"
	txtSelect = txtSelect & " PR Code " & v_txtPRCODE	
end if

	txtSQL = txtSQL & "ORDER BY hdr.fcba, hdr.prcode" 'ditambah ukie tgl 9 okt 2013 untuk Order by


if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

 'response.Write(txtsql)
response.Flush()
%>
	<body>
		<div id="wrapper">
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>PURCHASE REQUEST APPROVAL</h2>
				</div>
			</div>
					
		</th>
        </tr>
        		<!--<div class="printhide">-->
						<tr>
							<th colspan="11"  style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
						<tr>
							<%	
							
							Set rsCat = Server.CreateObject("ADODB.Recordset")
							rsCat.CursorType = 3
							rsCat.CursorLocation = 3
							rsCat.Open txtSQL, connstr

							
							if Not rsCat.eof then
							   nCountofField=rsCat.fields.count-1
							end if
							NO=1%>
						</tr>
						<tr>
							<!--<th>No.</th>-->
							<th>Date</th>
							<th>PR.Code</th>
							<th>Requestor</th>
							<th>Status</th>
							<th>Remarks</th>
							<th>Stock Code</th>
							<th>Stock Description</th>
							<th>Requested (Qty)</th>
							<th>Approved (Qty)</th>
							<th>Ordered (Qty)</th>
							<th>Next Approve</th>
						</tr>
                        </thead>
                        <tbody>
							<%
							Dim v_Opening,v_Debit,v_Credit,v_Saldo,v_Quantity
							v_Quantity=0
							v_Amount = 0
							While Not rsCat.EOF%>


							<tr>
								<!--<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%> </td>-->
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("prdate")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("prcode")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("RequestorId")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("RowstateId")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Remarks")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ItemCode")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%> </td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("requestedqty")),2)%> </td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("approvedqty")),2)%> </td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("orderedqty")),2)%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("NextApv")%> </td>
							</tr>
							  <%

							  rsCat.MoveNext
							  'response.Flush()
							  no=no+1
							Wend
							rsCat.close()
							rsCat=null%>
							<input  type="hidden" value="<%=txtTable%>" name="txtTable">
							<input  type="hidden" value="<%=txtModule%>" name="txtModule">
							</form>
                            </tbody>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	