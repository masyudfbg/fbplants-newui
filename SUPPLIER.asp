<%Option Explicit%>
<%
'On Error resume Next 
%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Supplier</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}

function CheckIsActive()
{
 var cIsActive = document.getElementById('ISACTIVE');
 if (cIsActive.checked != true )
	{
		alert("Is Active Belum di Click");
	}
}

function calcParam()
{
	var cIsActive = document.getElementById('ISACTIVE').checked;
	if(cIsActive)
	{
		document.getElementById("hide_ISACTIVE").value ="1";
	}else
	{
		document.getElementById("hide_ISACTIVE").value ="0";
	}
	
	var cIsBlock = document.getElementById('ISBLOCK').checked;
	if(cIsBlock)
	{
		document.getElementById("hide_ISBLOCK").value ="1";
	}else
	{
		document.getElementById("hide_ISBLOCK").value ="0";
	}
	
	var cIsTransporter = document.getElementById('ISTRANSPORTER').checked;
	if(cIsTransporter)
	{
		document.getElementById("hide_ISTRANSPORTER").value ="1";
	}else
	{
		document.getElementById("hide_ISTRANSPORTER").value ="0";
	}
	
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
 
<%

 '-------------awal----------------------- 

Dim FCENTRY, FCEDIT, FCIP, FCBA
Dim CUSTOMERCODE, txtDESCRIPTION, SUPPLIER, ICNO, BANK, BANKACCOUNTNO,BANK_OWNERSHIP, COMPANYNO,OWN_ESTATE,CUSTOMER_FCBA,CUSTFFBTYPE,ISTRANSPORTER
Dim cmdView,txt_CUSTOMERCODE
Dim vSTREET, vCITY, vPOSTALCODE, vPROVINCE, vTELEPHONE, vFAX, vEMAIL, vISACTIVE, vISBLOCK, vISTRANSPORTER
Dim hide_ISACTIVE, hide_ISBLOCK, hide_ISTRANSPORTER
txttable=request("txttable")


if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
 
	txt_CUSTOMERCODE=request("txt_CUSTOMERCODE")
	CUSTOMERCODE=request("CUSTOMERCODE")
	txtDESCRIPTION=request("txtDESCRIPTION")
	''SUPPLIER=request("SUPPLIER")
	ICNO=request("ICNO")
	BANK=request("BANK")
	BANKACCOUNTNO=request("BANKACCOUNTNO")
	BANK_OWNERSHIP=request("BANK_OWNERSHIP")
	COMPANYNO=request("COMPANYNO")

	OWN_ESTATE = "N" 'Default value
	SUPPLIER = "Y" 'Default value
	' ISTRANSPORTER = "N" 'Default value
	' ISTRANSPORTER = request("ISTRANSPORTER")
	
	OWN_ESTATE =request("OWN_ESTATE")
	CUSTOMER_FCBA =request("CUSTOMER_FCBA")
	CUSTFFBTYPE =request("CUSTFFBTYPE")

	vSTREET =request("STREET")
	vCITY =request("CITY")
	vPOSTALCODE =request("POSTALCODE")
	vPROVINCE =request("PROVINCE")
	vTELEPHONE =request("TELEPHONE")
	vFAX =request("FAX")
	vEMAIL =request("EMAIL")
	
	hide_ISACTIVE = 0
	hide_ISBLOCK = 0
	hide_ISTRANSPORTER = 0
	
	hide_ISACTIVE = request("hide_ISACTIVE") 
	hide_ISBLOCK = request("hide_ISBLOCK") 
	hide_ISTRANSPORTER = request("hide_ISTRANSPORTER") 
		
	vISACTIVE =request("ISACTIVE")
	vISBLOCK =request("ISBLOCK")
	vISTRANSPORTER =request("ISTRANSPORTER")
	

	if Request("txtJenis")= "Delete" then
		q	=	"DELETE FROM CUSTOMER WHERE  CUSTOMERCODE='" & CUSTOMERCODE & "' AND fcBA='" & Session("fcBA") & "'"
		connStr.execute(q)
		cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		q	=	" INSERT INTO CUSTOMER (" & _
				" CUSTOMERCODE, " & _
				" DESCRIPTION, " & _
				" SUPPLIER, " & _
				" ICNO, " & _
				" BANK, " & _
				" BANKACCOUNTNO, " & _
				" BANK_OWNERSHIP, " & _
				" COMPANYNO, " & _
				" OWN_ESTATE, " & _
				" CUSTOMER_FCBA, " & _
				" CUSTFFBTYPE, " & _
				" FCENTRY, " & _
				" FCEDIT, " & _
				" FCIP, " & _
				" FCBA, " & _
				" LASTUPDATE, " & _
				" LASTTIME, " & _
				" STREET, " & _
				" CITY, " & _
				" POSTALCODE, " & _
				" PROVINCE, " & _
				" TELEPHONE, " & _
				" FAX, " & _
				" EMAIL, " & _	
				" ISTRANSPORTER," & _				
				" ISACTIVE," & _
				" ISBLOCK " & _
				" ) values (" & _
				"   '" & CUSTOMERCODE & "'" & _			
				" , '" & txtDESCRIPTION & "'" & _
				" , '" & SUPPLIER & "'" & _
				" , '" & ICNO & "'" & _
				" , '" & BANK & "'" & _
				" , '" & BANKACCOUNTNO & "'" & _
				" , '" & BANK_OWNERSHIP & "'" & _
				" , '" & COMPANYNO & "'" & _
				" , '" & OWN_ESTATE & "'" & _
				" , '" & CUSTOMER_FCBA & "'" & _
				" , '" & CUSTFFBTYPE & "'" & _
				" , '" & Session("CodeName") & "'" & _					
				" , '" & Session("CodeName") & "'" & _
				" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _					
				" , '" & Session("fcBA") & "'" & _					
				" , sysdate " & _										
				" , To_Char(sysdate,'HH24:MI') " & _	
				" , '" & vSTREET & "'" & _
				" , '" & vCITY & "'" & _
				" , '" & vPOSTALCODE & "'" & _
				" , '" & vPROVINCE & "'" & _
				" , '" & vTELEPHONE & "'" & _
				" , '" & vFAX & "'" & _
				" , '" & vEMAIL & "'" & _
				" , '" & cdbl(0) & "'" & _				
				" , '" & cdbl(0) & "'" & _
				" , '" & cdbl(0) & "'" & _	
				" )"			 

		''response.write q
		'response.Flush()
		connStr.execute(q)
		cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q=	"  UPDATE CUSTOMER SET" & _ 
			"  DESCRIPTION ='" & TXTDESCRIPTION & "'" & _
			", ICNO ='" & ICNO & "'" & _
			", BANK ='" & BANK & "'" & _
			", BANKACCOUNTNO ='" & BANKACCOUNTNO &"'" & _
			", BANK_OWNERSHIP = '" & BANK_OWNERSHIP &"'" & _
			", COMPANYNO ='" & COMPANYNO & "'" & _
			", OWN_ESTATE ='" & OWN_ESTATE & "'" & _
			", CUSTOMER_FCBA ='" & CUSTOMER_FCBA & "'" & _
			", CUSTFFBTYPE ='" & CUSTFFBTYPE & "'" & _
			", SUPPLIER ='" & SUPPLIER & "'" & _
			", LASTUPDATE = sysdate " & _
			", LASTTIME = To_Char(sysdate,'HH24:MI') " & _
			", ISTRANSPORTER ='" & hide_ISTRANSPORTER & "'" & _
			", STREET ='" & vSTREET & "'" & _
			", CITY ='" & vCITY & "'" & _
			", POSTALCODE ='" & vPOSTALCODE & "'" & _
			", PROVINCE ='" & vPROVINCE & "'" & _
			", TELEPHONE ='" & vTELEPHONE & "'" & _
			", FAX ='" & vFAX & "'" & _
			", EMAIL ='" & vEMAIL & "'" & _	
			"  WHERE  CUSTOMERCODE='" & CUSTOMERCODE &"'"  & _
			"  AND fcBA='" & Session("fcBA") & "'" 
			
		connStr.execute(q)
		connStr.execute("commit")
		''response.Write(q)
		'response.Flush()
		''response.write q
		
		' ", ISTRANSPORTER ='" & hide_ISTRANSPORTER & "'" & _			
		' ", ISACTIVE ='" & hide_ISACTIVE & "'" & _
		' ", ISBLOCK ='" & hide_ISBLOCK & "'" & _	
			
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")

If Err.Number <> 0 then
   Response.Write "<font color=red>" & "Error Number : " & Err.Number &  " Desc. " & Err.Description & "</font>"
End if

 '-------------ISI Payment Form
  %>
  
<form name="frmEntry" method="post" action="SUPPLIER.asp">
	<table class="data" width="94%" cellpadding="0" cellspacing="0">
          <TR> 
            <td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE SUPPLIER INFORMATION</strong></font></div></td>
          </TR>
		<tr>
			<td align="left" width="15%">Supplier Code</td><td align="left"><input type="text" size="15" name="txt_CUSTOMERCODE" value="<%=request("txt_CUSTOMERCODE")%>"  ></td>
		</tr>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;
		<a href="PrintSupplier.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_CUSTOMERCODE=<%=REQUEST("txt_CUSTOMERCODE")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
		</td>
		</tr>
    </table>  
 </form>
 <%
 if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   ' set rsGAD = server.CreateObject("Adodb.recordset")
    'Field Numeric Jangan di Sum karena ditampilkan mesti level detail

	'' Query Menampilkan Data di Table Bawah  
	q	=	" SELECT " & _
			" CUSTOMERCODE, DESCRIPTION, ICNO, " & _
			" BANK, BANKACCOUNTNO,BANK_OWNERSHIP, COMPANYNO, OWN_ESTATE, CUSTOMER_FCBA, CUSTFFBTYPE, " & _			
			" STREET, CITY, TELEPHONE, ISTRANSPORTER, ISBLOCK, ISACTIVE,  " & _
			" POSTALCODE, PROVINCE, FAX, EMAIL " & _
			" FROM CUSTOMER " & _
			" WHERE fcba = '" & Session("fcBA") & "'" & _        
			" AND  UPPER(SUPPLIER) ='Y' "	
	 
    Set rsGAD = Server.CreateObject("ADODB.Recordset")
	' if  len(trim(request("CUSTOMERCODE"))) > 0 then
		' q = q & " and CUSTOMERCODE='" & request("CUSTOMERCODE") & "'"
	' end if
	
	''q = q
   
   ''response.Write(q)
   'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr 
   
  %>
   
   <!-- Form Entry And Edit --> 
   
   <form action="SUPPLIER.asp" method="post" name="eForm">
        <table class="data" cellpadding="0" cellspacing="0" width="100%" Border="0">
		
		<!--Garis Hijau /  Blok Greeen -->
		<tr bgcolor="#00a300"> 
			<td colspan="8">
				<div align="left">
					<font size="2" face="Verdana" color="#fff"><strong>Supplier</strong></font>
				</div>
			</td>
		</tr>
		
		<tr>
			<%IF Request("txtJenis")="Edit" then%>
				<td width="13%" align="left">Supplier Code</td>
				<td width="64%" align="left"><input type="text" size="15" maxlength="15" name="CUSTOMERCODE" value="<%=CUSTOMERCODE%>" readonly="True" style="background:#999999"><font color="#FF0000">*</font></td>
			<%else%>
				<td width="2%" align="left">Supplier Code</td>
				<td width="11%" align="left"><input type="text" size="15" maxlength="15" name="CUSTOMERCODE" value="<%=CUSTOMERCODE%>"  ><font color="#FF0000">*</font></td>
			<%end if%>
			
			<td align="left">Description</td>
			<td align="left" width="100%" ><input type="text" size="50" maxlength="255" name="txtDESCRIPTION" value="<%=txtDESCRIPTION%>"><font color="#FF0000">*</font></td>		
		</tr>
		
		<!--
		<tr>
			<%IF Request("txtJenis")="Edit" then%>
				<td align="left">Supplier</td>
				<td align="left"><select name="SUPPLIER">
					<%IF request("SUPPLIER")="Y" THEN %>
					<Option value="Y" selected>Yes</Option>
					<Option value="N">No</Option>
					<%ELSE%>
					<Option value="Y">Yes</Option>
					<Option value="N" selected>No</Option>
				  </select> </td>
					<%End If%>
			<%Else%>
				<td align="left">Supplier</td>
				<td align="left"><select name="SUPPLIER">
				<Option value="Y">Yes</Option>
				<Option value="N">No</Option>
				</select></td>
			<%End If%>
						
			<td align="left">ICNO</td>
			<td align="left" width="100%"><input type="text" size="50"  maxlength="240" name="ICNO" value="<%=ICNO%>"></td>
				
		</tr>
		-->
		
		<tr>
			<td align="left">Bank</td>
			<td align="left"><input type="text" size="15" maxlength="15" name="BANK" value="<%=BANK%>"></td>
			
			<td align="left">Bank Account No</td>
			<td align="left"><input type="text" size="40"  maxlength="40" name="BANKACCOUNTNO" value="<%=BANKACCOUNTNO%>"></td>
			<td align="left">Bank Ownership</td>
			<td align="left"><input type="text" size="40"  maxlength="40" name="BANK_OWNERSHIP" value="<%=BANK_OWNERSHIP%>"></td>
		</tr>
	
		<tr>
			<td align="left">Company No</td>
			<td align="left"><input type="text" size="30" maxlength="30" name="COMPANYNO" value="<%=COMPANYNO%>"></td>
						
			<td align="left">Own Estate</td>
			<td align="left">
				<select name="OWN_ESTATE" >
				<% 
				dim YES_,NO_
				if OWN_ESTATE = "Y" then
						YES_ = "selected"
					else
						NO_ = "selected"
					end if
				%>
					<option value="Y" <%=YES_%>>Yes</option>
					<option value="N" <%=NO_%>>No</option>
				</select>
			</td>
		</tr>

		<tr>
			<td align="left">Supplier FCBA</td>
			<td align="left"><input type="text" size="30" maxlength="30" name="CUSTOMER_FCBA" value="<%=CUSTOMER_FCBA%>"></td>
			
			<td align="left">Supplier FFB Type</td>
			<td align="left">
				<select name="CUSTFFBTYPE">
					<option VALUE="" ></option>
					<% 
						DIM cffbT,rscffbT,SELECTED
						cffbT="select FCCODE, FCNAME from CUSTOMERFFBTYPE"
						Set rscffbT = Server.CreateObject("ADODB.Recordset")
						'response.Write(cffbT)
						'response.Flush()
						rscffbT.CursorType = 3
						rscffbT.CursorLocation = 3
						rscffbT.Open cffbT, connstr 
						if Not rscffbT.eof then
						While Not rscffbT.EOF
						IF CUSTFFBTYPE = rscffbT("FCCODE") THEN
						SELECTED = "SELECTED"
						ELSE
						SELECTED = ""
						END IF
						%>
						<option VALUE="<%=rscffbT("FCCODE")%>" <%=SELECTED%>><%=rscffbT("FCNAME")%></option>
						<%
						rscffbT.MoveNext
						Wend
						end if 
						rscffbT.Close
						set rscffbT = Nothing
					%>
				</select>
			</td>
			
		</tr>
		
		<!--Garis Hijau /  Blok Greeen -->
		<tr bgcolor="#00a300"> 
			<td colspan="8">
				<div align="left">
					<font size="2" face="Verdana" color="#fff"><strong>Address</strong></font>
				</div>
			</td>
		</tr>
		
		<tr>			
			<td width="189" align="left">Street</td>		
			<td width="180" align="left"><input type="text" size="30" maxlength="30"  name="STREET" id="STREET" value="<%=vSTREET%>"></td>
			
			<td width="149" align="left">City</td>
			<td width="180" align="left"><input type="text" size="20" maxlength="30"  name="CITY" id="CITY" value="<%=vCITY%>"></td>
			
			<td width="149" align="left">PostalCode</td>
			<td width="180" align="left"><input type="text" size="20" maxlength="30"  name="POSTALCODE" id="POSTALCODE" value="<%=vPOSTALCODE%>"></td>
		</tr>
		<tr>
			<td width="189" align="left">Province</td>
			<td width="180" align="left"><input type="text" size="20" maxlength="30"  name="PROVINCE" id="PROVINCE" value="<%=vPROVINCE%>"></td>
			
			<td width="149" align="left">Telephone</td>
			<td width="180" align="left"><input type="text" size="20" maxlength="30"  name="TELEPHONE" id="TELEPHONE" value="<%=vTELEPHONE%>"></td>
			
			<td width="149" align="left">Fax</td>
			<td width="180" align="left"><input type="text" size="20" maxlength="30"  name="FAX" id="FAX" value="<%=vFAX%>"></td>
		</tr>
		<tr>
			<td width="189" align="left">Email</td>	
			<td width="180" align="left"><input type="text" size="30" maxlength="30"  name="EMAIL" id="EMAIL" value="<%=vEMAIL%>"></td>
		</tr>
		
		<!--Garis Hijau /  Blok Greeen -->
		<tr bgcolor="#00a300"> 
			<td colspan="8">
				<div align="left">
					<font size="2" face="Verdana" color="#fff"><strong>Information</strong></font>
				</div>
			</td>
		</tr>	

		<tr>
			<td width="149" align="left"></td>	
			<td width="180" align="left"> 
				<input type="hidden" name="hide_ISTRANSPORTER" id="hide_ISTRANSPORTER" value="<%=hide_ISTRANSPORTER%>">
				<input type="checkbox" size="10" maxlength="30" onclick="calcParam()"  name="ISTRANSPORTER" id="ISTRANSPORTER" value="<%if Request("hide_ISTRANSPORTER")="1" then%>" checked "<% end if %>" /> is Transporter
			</td>	
			
			<td width="180" align="left"> 
				<input type="hidden" name="hide_ISBLOCK" id="hide_ISBLOCK" value="<%=hide_ISBLOCK%>">
				<input type="checkbox" size="10" maxlength="30" disabled readonly onclick="calcParam()"  name="ISBLOCK" id="ISBLOCK" value="<%if Request("hide_ISBLOCK")="1" then%>" checked "<% end if %>" />  is Block
			</td>

			<td width="180" align="left">
				<input type="hidden" name="hide_ISACTIVE" id="hide_ISACTIVE" value="<%=hide_ISACTIVE%>">
				<input type="checkbox" size="10" maxlength="30" disabled readonly  onclick="calcParam()"  name="ISACTIVE" id="ISACTIVE" value="<%if Request("hide_ISACTIVE")="1" then%>" checked "<% end if %>"  /> is Active
			</td>			
		</tr>

		<tr> 
			<td width="189" align="left"></td>	
		</tr>
		<tr> 
			<td width="189" align="left"></td>	
		</tr>
		
		<tr>
			<td colspan="2" align="left">
				  <input type="Hidden" name="txt_CUSTOMERCODE" value="<%=REQUEST("CUSTOMERCODE")%>">
				<!--
				  <input type="text" name="hide_ISTRANSPORTER" value="<%=REQUEST("ISTRANSPORTER")%>">	
				  <input type="text" name="hide_ISBLOCK" 	   value="<%=REQUEST("ISBLOCK")%>">	
				  <input type="text" name="hide_ISACTIVE" 	   value="<%=REQUEST("ISACTIVE")%>">	
				-->
				
				  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				  <%IF Request("txtJenis")="Edit" then%>
					   <input type="hidden" name="txtJenis" value="Edit">
					   <input Type="Submit"  name="Add" value="Update">
				  <%Else%>
					   <input type="hidden" name="txtJenis" value="New">
					   <input Type="Submit"  name="Add" value="Add">
				  <%End if%>
				  <input Type="Reset"  name="Reset" value="Reset">
				  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
			</td>
		</tr>
   </table>
   </form>
  <!--#include file="validation_show.Asp"-->    
<%

	Dim intNumberOfPages,intRecords,nPage,howmanyfields,JumRec
	howmanyfields=rsGAD.fields.count-1
	
	JumRec=1
	intRecordsPerPage = 20
	
	rsGAD.PageSize = intRecordsPerPage
	intNumberOfPages = 0
	intNumberOfPages =  rsGAD.PageCount
	intRecords = rsGAD.RecordCount
	'Set our AbsolutePage property.
	if not rsGAD.EOF then
		If (isEmpty(Request("iPageNo")) or trim(Request("iPageNo"))="") Then
		  rsGAD.AbsolutePage = 1
		Else
		  rsGAD.AbsolutePage = Request("iPageNo")
		End If  
	end if
	
	if Request("iPageNo") = "" then
		nPage = 1
	Else
		nPage = Request("iPageNo")
	End if
	
   rsGAD.PageSize = intRecordsPerPage
   intNumberOfPages = 0
   intNumberOfPages =  rsGAD.PageCount
   intRecords = rsGAD.RecordCount   
   %>
   <table class="data" >
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD align="center" colspan="2">ACTION</TD>
	<TD align="left">SUPPLIER CODE</TD>	
	<TD align="left">DESCRIPTION</TD>	
	<!-- <TD align="left">SUPPLIER</TD>	 -->
	<TD align="left">ICNO</TD>	
	<TD align="left">BANK</TD>	
	<TD align="left">BANK ACCOUNT NO</TD>
	<TD align="left">BANK OWNERSHIP</TD>	
	<TD align="left">COMPANY NO</TD>
	<TD align="left">OWN ESTATE</TD>
	<TD align="left">SUPPLIER FCBA</TD>
	<TD align="left">SUPPLIER FFB TYPE</TD>	
	<TD align="left">STREET</TD>	
	<TD align="left">CITY</TD>	
	<TD align="left">TELEPHONE</TD>	
	<TD align="left">IS TRANSPORTER</TD>	
	<TD align="left">IS BLOCK</TD>	
	<TD align="left">IS ACTIVE</TD>	
	</tr>
   
   <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count - 5
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">
			<a href="SUPPLIER.asp?CUSTOMERCODE=<%=rsGAD("CUSTOMERCODE")%>&txttable=<%=request("txttable")%>&txtDESCRIPTION=<%=rsGAD("DESCRIPTION")%>&ICNO=<%=rsGAD("ICNO")%>&BANK=<%=rsGAD("BANK")%>&BANKACCOUNTNO=<%=rsGAD("BANKACCOUNTNO")%>&COMPANYNO=<%=rsGAD("COMPANYNO")%>&OWN_ESTATE=<%=rsGAD("OWN_ESTATE")%>&CUSTOMER_FCBA=<%=rsGAD("CUSTOMER_FCBA")%>&CUSTFFBTYPE=<%=rsGAD("CUSTFFBTYPE")%>&STREET=<%=rsGAD("STREET")%>&CITY=<%=rsGAD("CITY")%>&POSTALCODE=<%=rsGAD("POSTALCODE")%>&PROVINCE=<%=rsGAD("PROVINCE")%>&TELEPHONE=<%=rsGAD("TELEPHONE")%>&FAX=<%=rsGAD("FAX")%>&EMAIL=<%=rsGAD("EMAIL")%>&hide_ISTRANSPORTER=<%=rsGAD("ISTRANSPORTER")%>&hide_ISACTIVE=<%=rsGAD("ISACTIVE")%>&hide_ISBLOCK=<%=rsGAD("ISBLOCK")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>		
			</td>
			<td valign="middle" width="30">
			<a href="SUPPLIER.asp?CUSTOMERCODE=<%=rsGAD("CUSTOMERCODE")%>&txttable=<%=request("txttable")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
		    <%for i = 0 to nCountofField%>
			
			<% if rsGAD(i).Name="ISTRANSPORTER" OR rsGAD(i).Name="ISBLOCK" OR rsGAD(i).Name="ISACTIVE"  Then %>  <!--Filter Data Only Checkbox -->
				
				<% if rsGAD(i).Name="ISTRANSPORTER" AND rsGAD("ISTRANSPORTER")="1" Then%>  <!--  -->
					<td align="center" ><input type="checkbox" name="myTextEditBox" disabled readonly value="1" checked></td>
				<%Elseif rsGAD(i).Name="ISBLOCK" AND rsGAD("ISBLOCK")="1" Then%>
					<td align="center" ><input type="checkbox" name="myTextEditBox" disabled readonly value="1" checked></td>
				<%Elseif rsGAD(i).Name="ISACTIVE" AND rsGAD("ISACTIVE")="1" Then%>
					<td align="center" ><input type="checkbox" name="myTextEditBox" disabled readonly value="1" checked></td>
				<%Else%>
					<td align="center" ><input type="checkbox" name="myTextEditBox" disabled readonly value="0"></td>
				<%End If%>
				
	  		<%Else%>
				<td align="left" title="<%=rsGAD(i).Name%>" valign="top"><%=rsGAD(i)%></td>
			<%End If%>
				
			    
		    <%next%>
		  </tr>
		  <%
		  rsGAD.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsGAD.Close
   set rsGAD = Nothing%>
   </div>
   </TABLE>
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->
</body>
</html>
