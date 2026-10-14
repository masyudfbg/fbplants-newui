<%
'Option Explicit
%>
<%
'On Error resume Next
%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - View Stock Transfer Order </title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="Table_Html.css" type="text/css" /> <!-- For Hidden Column UNIQUE_KEY  ON TABLE HTML, Because Field UNIQUE_KEY Not View on Table HTML, Create BY AST 20180514 -->
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<script language="JavaScript" src="/js/date-picker.js"></script>  
<script language="JavaScript"> 
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}

/*
function SetCheckBox(CbCode, CbOriCode){    

    var StringAll;
	var value_;


	var CID = document.getElementById(CbOriCode).checked;
			
	if(document.getElementById(CbOriCode).checked == 1) {
		StringAll = "<input id='data_" + CbCode + "' name='data_" + CbCode + "' type='hidden' value='on'>";
	}
	else {
		StringAll = "<input id='data_" + CbCode + "' name='data_" + CbCode + "' type='hidden' value='off'>";
	}
	
    document.getElementById("div_" + CbCode ).innerHTML=StringAll;
}
*/

/*
function SetCheckBoxDefault(CbCode, CbOriCode){    

    var StringAll;
	var value_;


	var CID = document.getElementById(CbOriCode).checked;
			
	if(document.getElementById(CbOriCode).checked == 1) {
		StringAll = "<input id='data_default_" + CbCode + "' name='data_default_" + CbCode + "' type='hidden' value='on'>";
	}
	else {
		StringAll = "<input id='data_default_" + CbCode + "' name='data_default_" + CbCode + "' type='hidden' value='off'>";
	}
	
    document.getElementById("div_default_" + CbCode ).innerHTML=StringAll;
}
*/

/*
function checkAll(formname, checktoggle, vNilai)
{

	var StringCheck;
	var checkboxes = new Array(); 
		 
	checkboxes = document[formname].getElementsByClassName('ok');
	var myTab = document.getElementById('entry');
		
	for (var i=0; i <= checkboxes.length; i++)  {		
		var objCells = myTab.rows.item(i+1).cells; //Setting ID yang ingin di ambil dari kolom berapa

		var id_key = objCells.item(1).innerHTML;  //Pengambilan ID untuk filter ##Get ID_Key##
		var string1 = id_key;

		if (checkboxes[i].type == 'checkbox') {
			checkboxes[i].checked = checktoggle;
			
			if (checkboxes[i].checked == true) {
				StringCheck = "<input id='data_" + string1 + "' name='data_" + string1 + "' type='hidden' value='on'>";
			}else{
				StringCheck = "<input id='data_" + string1 + "' name='data_" + string1 + "' type='hidden' value='off'>";
			}
		}
			document.getElementsByClassName("test")[i].innerHTML=StringCheck;
	}
}
*/

</script> 
</head>
<body>
<!--#include file="top.jsp"-->
<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
		<% if Session("fcBa") <> "" then %>
         <!--#include file="menu.asp"-->
        <% end if %>
	</div>
    <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
    <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
    <div class="box_both" style="overflow:scroll;"> 
 <%
 '-------------awal-----------------------
 %>
 <%


Dim v_FCENTRY, v_FCEDIT, v_FCIP
DIM v_FCCODE
Dim cmdView
Dim defaultCompany, qCompanyAll, rsSqlstr
Dim tempCounterDefault, doneContinue
Dim rsModule,qModule,qInsert, rsModuleCheck
Dim txtFcba,v_FCBA,txtpono,v_TxtPONO
		

v_TxtSTONO = request("txtSTONO")
v_TxtDONO  = request("TxtDONO")
v_TxtPONO  = request("TxtPONO")

cmdView	= "True"

	qCompanyAll = " SELECT * FROM StockTransferOrder WHERE 1=1 "
				
	Set rsSqlstr = Server.CreateObject("ADODB.Recordset")
	rsSqlstr.Open qCompanyAll, connstr, adOpenStatic


txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")

'response.write v_TxtSTONO

Function CheckAvlStoDoRcv(v_TxtSTONO)
	Dim Result : Result = ""
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select BridgingSAP_PRPOGRN_API.CheckAvlStoDoRcv ('"& v_TxtSTONO &"') result" & _
					" from dual"
					
	'response.write strAvailable

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		Result = rsAvailable("result")
	Else
		Result = ""
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckAvlStoDoRcv = Result
End Function



If Request("txtJenis") =  "UnRelease" then
	    
		Msgbox_ =  CheckAvlStoDoRcv(v_TxtSTONO)
		
		'response.write Msgbox_

	     If Msgbox_ = "TRUE" Then	
            strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.UnReleaseSTO('" & v_TxtSTONO & "'); END; "
			'response.write strQueryStatus 
	        connStr.execute(strQueryStatus)
         else   		
 		    err.Number = 1001
		    err.description = "<br/>" & Msgbox_ 
		 End If
		 
        cmdView	= "True" 
End If		

if err.number <> 0 then
	connStr.RollbackTrans
    Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

 %>
 
 <form name="frmEntry" method="post" action="ViewStockTransferOrder.Asp">
	<table class="data" width="500">
		<h2>VIEW STOCK TRANSPORT ORDER </h2>
				
	        &nbsp;&nbsp;
		<tr>
			<td><font face="verdana" size="1">STO. NUMBER</font></td>
			<td>
				<select name="txtSTONO" >
					<option value="--ALL--">--ALL--</option>
				   <%
					Dim rsSTO, strSTO
					' strSTO =   " Select distinct BusinessUnitExt_Api.GetFcba(p.PLANT,p.STORLOC) Fcba " & _
								 ' "  from PurchaseOrder@iplas192l p " & _
								 ' "  where p.bridgestate=0 And p.Iplasstate=1  And  " & _
								 ' "  p.IsRelease=1   And p.potype='STO' and p.appsys='IPLAS'"
								 
					strSTO = " SELECT DISTINCT STONO FROM StockTransferOrder  WHERE 1=1  and fcBA = '" & session("fcBA") & "' Order By STONO "
					
					Set rsSTO = Server.CreateObject("ADODB.Recordset")
					rsSTO.Open strSTO, connstr, adOpenStatic
								
					if not rsSTO.eof then
						Do While Not rsSTO.eof
							if v_TxtSTONO = rsSTO("STONO") then %>
									<option value="<%=rsSTO("STONO")%>" Selected><%=rsSTO("STONO")%></option>
							<%else%>
									<option value="<%=rsSTO("STONO")%>"><%=rsSTO("STONO")%></option>
							<%end if			
						rsSTO.movenext
						loop
					end if
					rsSTO.Close
					set rsSTO = Nothing
					%>
				</select>
			</td>
			
		   <td><font face="verdana" size="1">DO. NUMBER</font></td>
		   <td>
				<select name="TxtDONO" >
					<option value="--ALL--">--ALL--</option>
					<%
						Dim rsDono, strDono
						' strDono =   " Select distinct p.Pono " & _
									 ' "  from PurchaseOrder@iplas192l p " & _
									 ' "  where p.bridgestate=0 And p.Iplasstate=1  And  " & _
									 ' "  p.IsRelease=1   And p.potype='STO' and p.appsys='IPLAS' order by p.Pono"

						strDono = " SELECT DISTINCT DONO FROM StockTransferOrder  WHERE 1=1  and fcBA = '" & session("fcBA") & "' Order By DONO "
						
						Set rsDono = Server.CreateObject("ADODB.Recordset")
						rsDono.Open strDono, connstr, adOpenStatic
									
						if not rsDono.eof then
							Do While Not rsDono.eof
								if v_TxtDONO = rsDono("DONO") then %>
										<option value="<%=rsDono("DONO")%>" Selected><%=rsDono("DONO")%></option>
								<%else%>
										<option value="<%=rsDono("DONO")%>"><%=rsDono("DONO")%></option>
								<%end if			
							rsDono.movenext
							loop
						end if
						rsDono.Close
						set rsDono = Nothing
					%>
				</select>
	       </td>
		   
		   <td><font face="verdana" size="1">PO. NUMBER</font></td>
		   		<td>
					<select name="TxtPONO" >
						<option value="--ALL--">--ALL--</option>
						<%
							Dim rspono, strpono
							' strpono =   " Select distinct p.Pono " & _
										 ' "  from PurchaseOrder@iplas192l p " & _
										 ' "  where p.bridgestate=0 And p.Iplasstate=1  And  " & _
										 ' "  p.IsRelease=1   And p.potype='STO' and p.appsys='IPLAS' order by p.Pono"
						
							strpono = " SELECT DISTINCT TRACKNO as Pono FROM StockTransferOrder  WHERE 1=1  and fcBA = '" & session("fcBA") & "' Order By TRACKNO "
						
							Set rspono = Server.CreateObject("ADODB.Recordset")
							rspono.Open strpono, connstr, adOpenStatic
										
							if not rspono.eof then
								Do While Not rspono.eof
									if v_TxtPONO = rspono("Pono") then %>
											<option value="<%=rspono("Pono")%>" Selected><%=rspono("Pono")%></option>
									<%else%>
											<option value="<%=rspono("Pono")%>"><%=rspono("Pono")%></option>
									<%end if			
								rspono.movenext
								loop
							end if
							rspono.Close
							set rspono = Nothing
						%>
					</select>
				</td>
			<td>
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="ViewData">
			</td>
		</tr>
    </table>  
 </form>
<%

   txtCekUsers="True"


if cmdView = "True" and txtCekUsers="True" Then 
   Dim rsVH,intRecordsPerPage,cColor
   set rsVH = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
  
	  
		 ' q = "  Select BusinessUnitExt_Api.GetFcba(p.PLANT,p.STORLOC) Fcba,p.* " & _
			 ' "  from PurchaseOrder@iplas192l p " & _
			 ' "  where p.bridgestate=0 And p.Iplasstate=1  And  " & _
			 ' "  p.IsRelease=1  And p.potype='STO' and p.appsys='IPLAS' " & whereCons
				 
		q = "  SELECT  ROWSTATEID,  " & _
			"  		FCBA, DONO, DELIVDATE, STONO, STODATE, " & _
			"  		DOCTYPE, DOCTYPEDESC, VENDOR, COMPANY, " & _
			"  		TRACKNO, TRACKLINENO, MATERIAL, MATERIALDESC, PLANT, STORLOC,  " & _
			"  		DOQTY, UOM, DORCVQTY, NETPRICE, FREZEE  " & _
			"  FROM StockTransferOrder " & _
			"  WHERE 1=1  and fcBA = '" & session("fcBA") & "'"
			
	if v_TxtSTONO = "" then 
	else
		If v_TxtSTONO <> "--ALL--" then
			q = q & " And  STONO = '" & v_TxtSTONO & "' "
		end if		
	end if
	
	if v_TxtDONO = "" then 
	else
		If v_TxtDONO <> "--ALL--" then
			q = q & " And  DONO = '" & v_TxtDONO & "' "
		end if
	end if
	
	if v_TxtPONO = "" then 
	else
		If v_TxtPONO <> "--ALL--" then
			q = q & " And  TRACKNO = '" & v_TxtPONO & "' "
		end if
	end if
	
   'response.write q
	
   Set rsVH = Server.CreateObject("ADODB.Recordset")
   rsVH.CursorType = 3
   rsVH.CursorLocation = 3
   rsVH.Open q, connstr
   
   %>
   <form action="ViewStockTransferOrder.Asp" method="post" name="eForm">
   <table id="entry" class="entry" cellpadding="1" cellspacing="1" width="620">
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;">

			<th align="center" valign="top">NO</th>
			<th align="center" valign="top" >ACTION</th>
			<th align="center" valign="top">STATUS</th>
			<th align="center" valign="top">FCBA</th>			
			<th align="center" valign="top">DELIVERY ORDER.NO</th>
			<th align="center" valign="top">DELIVERY DATE</th>
			<th align="center" valign="top">STO.NO</th>
			<th align="center" valign="top">STO.DATE</th>
			<th align="center" valign="top">DOC.TYPE</th>
			<th align="center" valign="top">DOC.TYPE DESC</th>
			<th align="center" valign="top">VENDOR</th>
			<th align="center" valign="top">COMPANY</th>
			<th align="center" valign="top">PO. No</th>
			<th align="center" valign="top">PO. Line No</th>
			<th align="center" valign="top">MATERIAL</th>
			<th align="center" valign="top">MATERIAL DESC</th>
			<th align="center" valign="top">PLANT</th>
			<th align="center" valign="top">STORLOC</th>
			<th align="center" valign="top">DO. QTY</th>
			<th align="center" valign="top">UOM</th>
			<th align="center" valign="top">DO RECEIVE QTY</th>
			<th align="center" valign="top">NET PRICE</th>
	</tr>
   <%
   if Not rsVH.eof then
      nCountofField=rsVH.fields.count-1
	  no=1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if

	%>
		  <tr bgcolor="<%=cColor%>">
		    <td align="left" valign="top"><%=NO%></td>
			<% If cint(rsVH("FREZEE")) = 0  Then %>
				<td align="left" valign="top">
				   <a href="ViewStockTransferOrder.Asp?txtSTONO=<%=rsVH("STONO")%>&fcba=<%=session("fcba")%>&txtJenis=UnRelease" onClick="return confirm('Are you sure to UnRelease STO NO :  <%=rsVH("STONO")%>  ?');">UnRelease</a>
				</td> 
			<% Else %>
			    <td align="left" valign="top">FREZEE</td> 
			<% End If %>
		    <%for i = 0 to nCountofField -1 %>
				<% if rsVH(i).Name="DELIVDATE" Then%>  <!--  -->
					<td align="left" title="<%=rsVH(i).Name%>" valign="top"><%=rsVH(i)%></td>
				<%Else%>
					<td align="left" title="<%=rsVH(i).Name%>" valign="top"><%=rsVH(i)%></td>
				<%End If%>
				
				<% 
					' response.write("Attr:" & rsVH(i).Attributes & "<br>")
					' response.write("Name:" & rsVH(i).Name & "<br>")
					' response.write("Value:" & rsVH(i).Value & "<br>")
					' response.write("Type:" & rsVH(i).type & "<br>""<br>")
				%>
		    <%next%>			  
			  			
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
	Else  
	%>
		<tr><td colspan="20" align="center" style="color:#FFFFFF">NO DATA AVAILABLE</td></tr>
	<%
		
   end if 
   rsVH.Close
   set rsVH = Nothing%>
	
		<input type="hidden" name="txtSTONO" value="<%=v_txtSTONO%>">
		<input type="hidden" name="txtDONO" value="<%=v_txtDONO%>">
		<input type="hidden" name="txtPONO" value="<%=v_txtPONO%>">
		
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input type="hidden" name="txtJenis" value="Edit">
	  
   </TABLE>
   </div>
   </form>
<%elseif cmdView = "True" and txtCekRoles="False" Then %>
   <P><font face="Verdana" size="1" color="#FF0000">User Name Not Found.. Please entry from Data Setup</P>
<%End if%>  
 
  <%
  '-------------akhir----------------------
  %>
</div></div>
  <!--#include file="bottom.jsp"-->

  <!-- Additional : AST, Date: 18 May 2018 ( 2 Ramadhan 1439 H ), Function	: For paging Table Html every Ten Data -->
  <script type="text/javascript" src="/js/Paging_TableHtml.js"></script>  
  
</body>
</html>
