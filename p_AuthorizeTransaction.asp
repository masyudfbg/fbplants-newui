<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Purchase Request Transaction</title>
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
v_txtpurchaseby=trim(request("ManagedId"))
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

txtSQL = 	"  SELECT purchaserequest.prcode AS PRCODE," & _
			"         purchaserequest.prdate AS PRDATE," & _
			"         purchaserequest.managedid AS managedid," & _
			"         purchaserequestdetails.itemcode AS ITEMCODE," & _
			"         stock.fcname AS ITEMNAME," & _
			"         purchaserequestdetails.requestedqty AS REQUESTEDQTY," & _
			"         purchaserequestdetails.approvedqty AS APPROVEDQTY," & _
			"         get_auth_lastlevel (purchaserequest.fcba, purchaserequest.prcode) LASTAUTHORIZERLEVELCODE," & _
			"         authorizelevel.fcname AS AUTHORIZELEVELNAME," & _
			"         get_auth_lastleveluser (purchaserequest.fcba, purchaserequest.prcode) LASTAUTHORIZER," & _
			"         purchaserequestdetails_hist.lastupdate AS LASTACTIONDATE," & _
			"         get_auth_nextlevel (purchaserequest.fcba, purchaserequest.prcode) NEXTAUTHORIZERLEVELCODE," & _
			"         get_auth_nextleveluser (purchaserequest.fcba, purchaserequest.prcode) NEXTAUTHORIZERUSERCODE" & _
			"    FROM purchaserequestdetails_hist" & _
			"         INNER JOIN purchaserequestdetails" & _
			"            ON ( purchaserequestdetails_hist.fcba = purchaserequestdetails.fcba AND purchaserequestdetails_hist.prcode = purchaserequestdetails.prcode)" & _
			"         INNER JOIN purchaserequest" & _
			"            ON ( purchaserequestdetails.fcba = purchaserequest.fcba AND purchaserequestdetails.prcode = purchaserequest.prcode)" & _
			"         INNER JOIN stock" & _
			"            ON ( stock.fccode = purchaserequestdetails.itemcode AND stock.fcba = purchaserequestdetails.fcba)" & _
			"         INNER JOIN authorizelevel" & _
			"            ON ( purchaserequestdetails_hist.currentauthorizelevel = authorizelevel.fccode)" & _
			"         INNER JOIN" & _
			"         (  SELECT purchaserequestdetails_hist.prcode," & _
			"                   MAX (purchaserequestdetails_hist.currentauthorizelevel) alevel" & _
			"              FROM purchaserequestdetails_hist" & _
			"          GROUP BY purchaserequestdetails_hist.prcode) aa" & _
			"            ON ( purchaserequestdetails_hist.currentauthorizelevel = aa.alevel AND purchaserequestdetails_hist.prcode = aa.prcode)" & _
			"   WHERE     purchaserequest.fcba = '"& businessunit &"'" 
			
if  len(trim(v_txtPRDATE_From)) > 0 then
    txtSQL = txtSQL & " and purchaserequest.prdate >=to_Date('" & v_txtPRDATE_From & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " From " & FormatDateTime(v_txtPRDATE_From,1)	
end if
if  len(trim(v_txtPRDATE_To)) > 0 then
    txtSQL =txtSQL &  " and purchaserequest.prdate <=to_Date('" & v_txtPRDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " To " & FormatDateTime(v_txtPRDATE_To,1)	
end if
if  len(trim(v_txtPRCODE)) > 0 then
    txtSQL = txtSQL & " and purchaserequest.prcode='" & v_txtPRCODE & "'"
	txtSelect = txtSelect & " PR Code " & v_txtPRCODE	
end if
if  len(trim(v_txtpurchaseby)) > 0 then
    txtSQL = txtSQL & " and purchaserequest.managedid='" & v_txtpurchaseby & "'"
	txtSelect = txtSelect & " Purchase By " & v_txtpurchaseby	
end if
	txtSQL = txtSQL &	" GROUP BY purchaserequest.prcode," & _
						"         purchaserequest.fcba," & _
						"         purchaserequest.prdate," & _
						"         purchaserequest.managedid," & _
						"         purchaserequestdetails.itemcode," & _
						"         purchaserequestdetails.requestedqty," & _
						"         purchaserequestdetails.approvedqty," & _
						"         authorizelevel.fccode," & _
						"         authorizelevel.fcname," & _
						"         purchaserequestdetails_hist.fcedit," & _
						"         purchaserequestdetails_hist.lastupdate," & _
						"         stock.fcname" 
 
	txtSQL = txtSQL & " ORDER BY purchaserequest.prcode, purchaserequestdetails.itemcode" 


if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

' response.Write(txtsql)
response.Flush()
%>
	<body>
		<div id="wrapper">
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>PURCHASE REQUEST TRANSACTION</h2>
				</div>
			</div>
					
		</th>
        </tr>
        		<!--<div class="printhide">-->
						<tr>
							<th colspan="13"  style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
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
							<th>PR Code</th>
							<th>Purchase By</th>
							<th>Date</th>
							<th>Item Code</th>
							<th>Item Name</th>
							<th>Requested Quantity</th>
							<th>Approved Quantity</th>
							<th>Last Authorize Level Code</th>
							<th>Last Authorizer level Name</th>
							<th>Last Authorizer</th>
							<th>Last Action Date</th>
							<th>Next Level Authorizer</th>
							<th>Next Authorizer User</th>
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
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("prcode")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("managedid")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("prdate")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("itemcode")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("itemname")%> </td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("requestedqty"))),2)%> </td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("approvedqty"))),2)%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LASTAUTHORIZERLEVELCODE")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("authorizelevelname")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("lastauthorizer")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("lastactiondate")%> </td>
								<% if rscat("authorizelevelname") <> "FINAL PASS" then %>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("NEXTAUTHORIZERLEVELCODE")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("NEXTAUTHORIZERUSERCODE")%> </td>
								<%else%>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1">Approval Has Completed</td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1">Approval Has Completed</td>
								<% end if %>
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