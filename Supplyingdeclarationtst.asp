<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Supplying Declaration</title>

<link rel="StyleSheet" href="dtree.css" type="text/css" />
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script>-->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>

<script>
		function hanyaAngka(evt) {
		  var charCode = (evt.which) ? evt.which : event.keyCode
		   if (charCode > 31 && (charCode < 48 || charCode > 57))
 
		    return false;
		  return true;
		}
	</script>
</head>
<body>
<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">
 <!-------------awal!------------->
 <%

Dim qSupplyingDeclaration, rsSupplyingDeclarationInfo, tempQuantity, tempDecType
Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA, v_FCCODE, v_FCNAME
Dim qSupplyingInfo, rsSupplyingInfo, qDeclarationType, rsDeclarationType
Dim DECLARATIONTYPE, parcode, totalstand
Dim FIELDCODE 
Dim QUANTITY
Dim REMARKS
Dim TRANSDATE
Dim cmdView
Dim txtDocumentNo ,txtlineno
Dim txtmonth, txtyear
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

declarationtype =request("declarationtype")
TRANSDATE =request("TRANSDATE")
quantity=request("quantity")
remarks=request("remarks")
fieldcode=request("fieldcode")
txtDocumentNo = Request("txtDocumentNo")
txtlineno = Request("txtlineno")
If Len(Trim(TRANSDATE)) = 0 Then TRANSDATE = textDate(date)

txtyear = year(TRANSDATE)
txtmonth = Month(TRANSDATE)

Function CheckFieldCode(initialFieldCode, initialFcba, initialDate)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total" & _
	                " from dual " & _
					" where field_api.GetValid_Activation('" & initialFieldCode & "', '" & initialFcba & "', to_date('"& initialDate &"', 'mm/dd/yyyy')) = 'Y'"
	' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckFieldCode = initialAvailable
End Function

' Function CheckQtySupplying(initialFieldCode, initialFcba, initYear, initMonth, initquantity)
' 	Dim initialAvailable
' 	Dim rsAvailable, strAvailable

' 	strAvailable =	" select count(1) total" & _
' 	                " from supplyingprogram " & _
' 					" where fieldcode ='" & initialFieldCode & "'" & _
' 					" and year = '" & initYear & "'" & _
' 					" and quantity = '" & initquantity & "'" & _
' 					" and fcba = '" & initialFcba & "'"

' 	 ' response.write stravailable
' 	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
' 	rsAvailable.Open strAvailable, connstr, 3, 3
	
' 	If Not rsAvailable.EOF Then
' 		initialAvailable = Cdbl(rsAvailable("Total"))
' 	Else
' 		initialAvailable = 0
' 	End If
' 	strAvailable = ""
' 	rsAvailable.Close
' 	set rsAvailable = Nothing

' 	CheckQtySupplying = initialAvailable
' End Function

Function CheckQtyPartial(initfieldcode, initialFcba, inityear)
	Dim initialAvailable
	Dim rsAvailable, strAvailable

	strAvailable =	" SELECT sum(sdd.quantity) total " & _
		  " FROM supplyingdeclarationdtl sdd" & _
		  " INNER JOIN supplyingdeclarationhdr sdh"& _
		  " ON (sdd.documentno = sdh.documentno AND sdd.fcba = sdh.fcba)" & _
		  " WHERE sdd.fcba = '" & initialfcba & "'" & _
		  " AND sdd.fieldcode = '" & initfieldcode & "'" & _
		  " AND to_char(sdh.transdate,'yyyy') = '" & inityear & "'"

	 ' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(to_nol(rsAvailable("Total")))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckQtyPartial = initialAvailable
End Function

Function CheckPlantingDate(initFieldCode, initFcba, initDate)
	Dim initialAvail
	Dim rsAvail, strAvail
	
	strAvail =	" select count (1) total" & _
	            " from dual " & _
	            " where to_char(field_api.GetValid_Plantingdate('" & initFieldCode & "', '" & initFcba & "', to_date('"& initDate &"', 'mm/dd/yyyy')),'yyyy')= to_char(to_date('"& initDate &"', 'mm/dd/yyyy'),'yyyy')"
	' response.write stravail
	Set rsAvail = Server.CreateObject("ADODB.Recordset")
	rsAvail.Open strAvail, connstr, 3, 3
	
	If Not rsAvail.EOF Then
		initialAvail = Cdbl(rsAvail("Total"))
	Else
		initialAvail = 0
	End If
	strAvail = ""
	rsAvail.Close
	set rsAvail = Nothing

	CheckPlantingDate = initialAvail
End Function

Function F_CurrentTotalStand(initFieldCode, initFcba, initDate)
	Dim initialAvail
	Dim rsAvail, strAvail
	
	strAvail =	" select field_api.getValid_POKOK('" & initFieldCode & "', '" & initFcba & "', to_date('"& initDate &"', 'mm/dd/yyyy')) total" & _
	            " from dual " 
	' response.write stravail
	Set rsAvail = Server.CreateObject("ADODB.Recordset")
	rsAvail.Open strAvail, connstr, 3, 3
	
	If Not rsAvail.EOF Then
		initialAvail = Cdbl(rsAvail("Total"))
	Else
		initialAvail = 0
	End If
	strAvail = ""
	rsAvail.Close
	set rsAvail = Nothing

	F_CurrentTotalStand = initialAvail
End Function

If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
	    
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), TRANSDATE)
		If Cek_IsDailyClosing = "FALSE" then
			q = " delete from supplyingdeclarationhdr " & _
				" 		where documentno = '" & txtDocumentNo & "'" & _
				"       and fcBA = '" & Session("fcBA") & "'"
							' Response.Write q
							connStr.execute(q)
			
			if err.number = 0 then
			   connStr.CommitTrans
			end if
		else
			err.number = 1001
			err.description = "Tanggal """ & TRANSDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
        cmdView	= "True"
	elseif Request("txtJenis")= "Delete" then

			

			q = " delete from supplyingdeclarationdtl " & _
						" 		where documentno = '" & txtDocumentNo & "'" & _
						"       and lineno = '" & txtlineno & "'" & _
						"       and fcBA = '" & Session("fcBA") & "'"
				' Response.Write q
				connStr.execute(q)

					cmdView	= "True"

	elseif Request("txtJenis")="New" then		
			Dim rsSupplyingDeclarationHdr
			Set rsSupplyingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
			rsSupplyingDeclarationHdr.Open " select documentno from supplyingdeclarationhdr where documentno = '" & txtDocumentNo & "' and fcba ='" & Session("fcba") & "' ", connstr
			If rsSupplyingDeclarationHdr.EOF Then
				q = " insert into supplyingdeclarationhdr" & _
					" (documentno, transdate, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
					" '" &txtDocumentNo & "'" & _
					" , to_date('" & TRANSDATE & "','mm/dd/yyyy')" & _ 
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" , '" & Session("fcba") & "'" & _
					" , sysdate" & _
					" , to_char(sysdate, 'hh24:mi')" & _
					" )"								
				'Response.Write q
				connStr.execute(q)
				
			End If
			rsSupplyingDeclarationHdr.Close
			set rsSupplyingDeclarationHdr = Nothing
					
			If Len(Trim(fieldcode)) = 0 Then
				Err.Number = 1001
				Err.Description = "Fieldcode Harus diisi, tidak boleh kosong"
			Else
				Dim val_SupplyingCode : val_SupplyingCode = 0
				val_SupplyingCode = CheckFieldCode(fieldcode, session("fcBA"), TRANSDATE)
			If val_SupplyingCode = 0 Then
				Err.Number = 1001
				Err.Description = "Field Code """ & fieldcode & """ tidak ada di master Field iPlaS, Koreksi kembali Kode Blok yang anda masukkan."
			Else
			' 	Dim val_QtySupplying : val_QtySupplying = 0
			' 	val_QtySupplying = CheckQtySupplying(fieldcode, session("fcBA"), txtyear, txtmonth, quantity)
			' If cstr(""&declarationtype) = cstr("3")  Then
			' 	Err.Number = 1001
			' 	Err.Description = "Quantity Total Stand melebihi Quantity Supplying Program, Koreksi kembali Quantity Total Stand yang anda masukkan."
			' Else
			Dim val_QtySupplying : val_QtySupplying = 0
			val_QtySupplying = CheckQtyPartial(fieldcode, session("fcBA"), year(TRANSDATE))
				strFieldDetail =	"	select sum(quantity) quantity" & _
									"	from supplyingprogram " & _
									" 	where year = '" & txtyear & "'" & _
									"   and fieldcode = '" & fieldcode & "'" & _
									"   and fcBA = '" & Session("fcBA") & "'"								
									' Response.Write strFieldDetail
									Set rsQty = Server.CreateObject("ADODB.Recordset")
									rsQty.Open strFieldDetail, connstr							
				if not rsQty.Eof then
				do while not rsQty.eof
						maxquantity = cdbl (to_nol(rsQty("quantity")))
						rsQty.MoveNext
					loop
				End If
				If cdbl (to_nol(quantity)) + val_QtySupplying > cdbl (to_nol(maxquantity)) Then
					Err.Number = 1001
					Err.Description = "Anda melebihi Quantity Total Stand pada Supplying Program, mohon cek kembali data yang anda masukan"
				Else

				qLineMax = " select nvl(max(lineno),0) tempMaxLine" & _
											   " from supplyingdeclarationdtl" & _
											   " where documentno = '" & txtDocumentNo & "'" & _
											   " and fcba = '" & session("fcBA") & "'"			
				set rsqLineMax = Server.CreateObject("ADODB.Recordset")
				rsqLineMax.Open qLineMax, connstr
				if not rsqLineMax.eof then
					tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
				else
					tempMaxLine = 0
				end if 

				tempMaxLine = tempMaxLine + 1
				txtLineNoHist = tempMaxLine

				Dim val_Plantingdate : val_Plantingdate = 0
				val_Plantingdate = CheckPlantingDate(fieldcode, session("fcBA"), TRANSDATE)
				Current_TotalStand = F_CurrentTotalStand(fieldcode, session("fcBA"), TRANSDATE)
					
					If QUANTITY <= 0 Then
						err.number = 1001
						Err.Description = "Anda harus Masukkan Angka di atas 0"
					Else
						q = " insert into SUPPLYINGDECLARATIONDTL" & _
							" (declarationtype, fieldcode, quantity, remarks, documentno, lineno, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
							" '" & declarationtype & "'" & _
							" , '" & fieldcode & "'" & _
							" , '" & quantity & "'" & _
							" , '" & remarks & "'" & _
							" , '" & txtDocumentNo & "'" & _
							" , '" & tempMaxLine & "'" & _
							" , '" & Session("CodeName") & "'" & _
							" , '" & Session("CodeName") & "'" & _
							" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
							" , '" & Session("fcba") & "'" & _
							" , sysdate" & _
							" , to_char(sysdate, 'hh24:mi')" & _
							" )"
							' Response.Write q
							connStr.execute(q)
					' End If
				End If
			End If
		End If
	End If
	cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
				If Len(Trim(fieldcode)) = 0 Then
					Err.Number = 1001
					Err.Description = "Fieldcode Harus diisi, tidak boleh kosong"
				Else
					Dim val_SupplyingCodes : val_SupplyingCodes = 0
					val_SupplyingCodes = CheckFieldCode(fieldcode, session("fcBA"), TRANSDATE)
				If val_SupplyingCodes = 0 Then
					Err.Number = 1001
					Err.Description = "Field Code """ & fieldcode & """ tidak ada di master Field iPlaS, Koreksi kembali Kode Blok yang anda masukkan."
				Else
					
				Dim val_Plantingdates : val_Plantingdates = 0
				val_Plantingdates = CheckPlantingDate(fieldcode, session("fcBA"), TRANSDATE)
				Current_TotalStand = F_CurrentTotalStand(fieldcode, session("fcBA"), TRANSDATE)
				
				' If cstr(""&declarationtype) = cstr("1")  Then
					If QUANTITY <= 0 Then
						err.number = 1001
						Err.Description = "Anda harus Masukkan Angka di atas 0"
					
				Else
					q =	" Update supplyingdeclarationdtl" & _
						" set quantity = '" & quantity & "'" & _
						"     , remarks = '" & remarks & "'" & _
						"     , fcentry = '" & Session("CodeName") & "'" & _
						"     , fcedit = '" & Session("CodeName") & "'" & _
						" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"     , lastupdate = sysdate " & _
						"     , lasttime = to_char(sysdate,'hh24:mi') " & _
						" 		where documentno = '" & txtDocumentNo & "'" & _
						"       and lineno = '" & txtlineno & "'" & _
						"       and fcBA = '" & Session("fcBA") & "'"
		 			' response.write q
					connStr.execute(q)
					
					End If
				End If
			End If
		End If
	End If
	cmdView	= "True"	
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

Dim rsSupplyingDeclaration, strSuppyingDeclaration
strSuppyingDeclaration	=	" select documentno, to_char(TRANSDATE,'mm/dd/yyyy') TRANSDATE" & _
							" from supplyingdeclarationhdr" & _
							" where documentno = '" & txtDocumentNo & "'" & _
							" and fcba = '" & session("fcBA") & "'"
						
set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
if Not rsSupplyingDeclaration.Eof then
	txtDocumentNo	=	rsSupplyingDeclaration("documentno")
	TRANSDATE	=	rsSupplyingDeclaration("TRANSDATE")	
End if
'response.Write strSuppyingDeclaration
rsSupplyingDeclaration.Close
Set rsSupplyingDeclaration = Nothing

If Err.Number <> 0 then
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

Dim val_IsTrue : val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & TRANSDATE & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & TRANSDATE & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('costbook')"
			
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if

'response.Write strNurseryCulling
oRecordSet.Close
Set oRecordSet = Nothing
%>

 <!-------------ISI Supplying Declaration Form!------------->
 
	<form name="frmEntry" method="post" action="Supplyingdeclarationtst.Asp">
		<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">Document No</td>
			<td align="left">
				<input type="text" maxlength="15" name="txtDocumentNo" size="30" value="<%=txtDocumentNo%>" autocomplete="Off">
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Document Date</td>
			<td align="left"><input class="datepicker" type="text" size="15" name="TRANSDATE" value="<%=TRANSDATE%>"></td>
		</tr>
		
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				
				<%If val_IsTrue = "Y" Then%>
					<a href="Supplyingdeclarationtst.Asp?txtDocumentNo=<%=txtDocumentNo%>&TRANSDATE=<%=TRANSDATE%>&txtJenis=HeaderDelete">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=txtDocumentNo%>" border="0" onClick="return confirm('Are you sure to delete');">
					</a>
				<% else %>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"> 
				<%end if%>
				<!-- HADIR -->
				<!-- <%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Nursery"
				dataSubmodule 	= "Culling"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a> -->
				<!-- END:HADIR -->
			</td>
		</tr>
		</table>
	</form>
<%
if cmdView = "True" and len(trim(txtDocumentNo)) > 0 Then
	Dim rsVH, intRecordsPerPage, cColor
	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	= " SELECT sdd.documentno, sdh.transdate, sdd.declarationtype, sdd.fieldcode, sdd.quantity, sdd.remarks, sdd.lineno ,pr.FCNAME as declarationtypename" & _
		  " FROM supplyingdeclarationdtl sdd" & _
		  " INNER JOIN supplyingdeclarationhdr sdh"& _
		  " ON (sdd.documentno = sdh.documentno AND sdd.fcba = sdh.fcba)" & _
		  " INNER JOIN parameter pr ON (sdd.fcba = pr.fcba AND sdd.declarationtype = pr.fccode AND pr.parcode = 'AGROTECH_SUPPLYINGTYPE')" & _
		  " WHERE sdd.fcba = '" & session("fcBA") & "'" & _
		  " AND sdd.documentno = '" & txtDocumentNo & "'"

	'response.Write q
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
%>
	<form action="Supplyingdeclarationtst.Asp" method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="100%">
			<tr> 
	            <td align="left"><font face="Verdana" size="1">Type</font></td>
					<td align="left"> 
						<select name="DECLARATIONTYPE" required="required" style="background:#FFFFCC;" title="DECLARATION TYPE Character a s/d z or 1 s/d 9">
							<option value=""></option> 
								<% 	Dim rsComboV
								q = "SELECT * FROM PARAMETER WHERE PARCODE = 'AGROTECH_SUPPLYINGTYPE' and fcba = '" & Session("fcba") & "'"
								Set rsComboV = Server.CreateObject("ADODB.Recordset")
								rsComboV.CursorType = 3
								rsComboV.CursorLocation = 3
								rsComboV.Open q, connstr
								if not rsComboV.eof then%>
								<%
								Do While Not rsCombov.eof
								%>
									<option value="<%=rsComboV("FCCODE")%>" 
										<%if cstr(""&DECLARATIONTYPE)=cstr(""&rsComboV("FCCODE")) then %> selected 
										<%End If%>><%=rsComboV("FCNAME")%>
									</option> 
								<%rsComboV.movenext
								loop
								End if
								rsComboV.Close
								set rsComboV = Nothing
								q=""
								%>
						</select> <font color="#FF0000">*</font>
					</td>
          		</tr>
			<tr>
				<td align="left" width="20%">Field Code</td>
				<td align="left">
					<input type="text" size="15" name="FIELDCODE" style="background:#FFFFCC;"  value="<%=fieldcode%>" maxlength="32" autocomplete="off">

					<font color="#FF0000">*</font>
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Quantity</td>
				<td align="left">
					<input type="text" size="15" name="QUANTITY" value="<%=quantity%>" autocomplete="off" maxlength="32" onkeypress="return hanyaAngka(event)">
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Remarks</td>
				<td align="left">
					<textarea name="REMARKS" cols="50" rows="3" maxlength="400" autocomplete="off" ><%=remarks%></textarea>
				</td>
			</tr>
			<tr>
				<td align="left" colspan="2">
					<input type="hidden" name="txtDocumentNo" value="<%=txtDocumentNo%>">
					<input type="hidden" name="TRANSDATE" value="<%=TRANSDATE%>">
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<%If val_IsTrue = "Y" Then%>
					<%IF Request("txtJenis")="Edit" then%>
						<input type="hidden" name="txtJenis" value="Edit">
						<input type="hidden" name="txtlineno" value="<%=txtlineno%>">
						<input Type="Submit"  name="Add" value="Update">
					<%Else%>
						<input type="hidden" name="txtJenis" value="New">
						<input Type="Submit"  name="Add" value="Add">
					<%End if%>
					<input Type="Reset"  name="Reset" value="Reset">
					<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
				<%End if%>
				</td>
			</tr>
		</table>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="left">NO</td>
		<td colspan="1" >&nbsp;</td>
		<td align="left">DOCUMENT NO</td>
		<td align="left">DOCUMENT DATE</td>
		<td align="left">TYPE</td>
		<td align="left">FIELDCODE</td>
		<td align="left">QUANTITY</td>
		<td align="left">REMARKS</td>
	</tr>
   
   <%
	if Not rsVH.eof then
		nCountofField = rsVH.fields.count - 1
		no = 1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
			<td align="left" valign="top"><%=NO%></td>
			
			
			<td valign="middle" width="50" >
				<a href="Supplyingdeclarationtst.Asp?txtDocumentNo=<%=txtDocumentNo%>&transdate=<%=rsVH("TRANSDATE")%>&declarationtype=<%=rsVH("declarationtype")%>&fieldcode=<%=rsVH("fieldcode")%>&quantity=<%=rsVH("quantity")%>&txtlineno=<%=rsVH("lineno")%>&Remarks=<%=rsVH("remarks")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("txtDocumentNo")%>" border="0"></a>
				
				<a href="Supplyingdeclarationtst.Asp?txtDocumentNo=<%=txtDocumentNo%>&transdate=<%=rsVH("TRANSDATE")%>&declarationtype=<%=rsVH("declarationtype")%>&fieldcode=<%=rsVH("fieldcode")%>&txtlineno=<%=rsVH("lineno")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtDocumentNo")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
				
			</td>

			    <td align="left" title="<%=rsVH("documentno").Name%>" valign="top"><%=rsVH("documentno")%></td>
			    <td align="left" title="<%=rsVH("transdate").Name%>" valign="top"><%=rsVH("transdate")%></td>
			    <td align="left" title="<%=rsVH("declarationtypename").Name%>" valign="top"><%=rsVH("declarationtypename")%></td>
			    <td align="left" title="<%=rsVH("fieldcode").Name%>" valign="top"><%=rsVH("fieldcode")%></td>
			    <td align="left" title="<%=rsVH("quantity").Name%>" valign="top"><%=rsVH("quantity")%></td>
			    <td align="left" title="<%=rsVH("remarks").Name%>" valign="top"><%=rsVH("remarks")%></td>
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
