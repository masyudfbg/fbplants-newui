<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - FFB Price List</title>
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

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
 <%'-------------awal-----------------------
 %>
 <%


Dim FCENTRY, FCEDIT, FCIP, FCBA
Dim PRODUCTCODE, ONDATE, PRICE,PLANTINGYEAR
dim txt_SHIFTCODE,Perioddate
Dim PRODUCTIONDATE, SHIFT_SHIFTCODE, ACTUALSTART, ACTUALEND, CAGESPROCESSED, CAGEWEIGHT, RAMPBALANCE, CAGESBALANCE, PLATFORMBALANCE, DOWNTIME, CAUSE, TOTALPRODUCTIONHOUR
Dim cmdView
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
txttable=request("txttable")

ONDATE =request("ONDATE")
PRODUCTCODE=request("ProductCODE")
PRICE=request("PRICE")
PRODUCTIONDATE=request("PRODUCTIONDATE")
PLANTINGYEAR	= Request("PLANTINGYEAR")

Perioddate=request("Perioddate")
txt_SHIFTCODE=request("txt_SHIFTCODE")
SHIFT_SHIFTCODE=request("SHIFT_SHIFTCODE")
ACTUALSTART=request("ACTUALSTART")
ACTUALEND=request("ACTUALEND")
CAGESPROCESSED=request("CAGESPROCESSED")
CAGEWEIGHT=request("CAGEWEIGHT")
RAMPBALANCE=request("RAMPBALANCE")
CAGESBALANCE=request("CAGESBALANCE")
PLATFORMBALANCE=request("PLATFORMBALANCE")
DOWNTIME=request("DOWNTIME")
CAUSE=request("CAUSE")
TOTALPRODUCTIONHOUR = request("TOTALPRODUCTIONHOUR")

 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then

	if Request("txtJenis")= "Delete" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), PRODUCTIONDATE)
		If Cek_IsDailyClosing = "FALSE" then
		   q="Delete from PriceListTBS " & _
		    " where" & _
		    " PRODUCTCODE = '" & PRODUCTCODE & "'" & _
			" And ONDATE = to_date('" & ONDATE & "','mm/dd/yyyy')" & _
			" AND PLANTINGYEAR = "  & PLANTINGYEAR & " " & _
			" and fcBA = '" & Session("fcBA") & "'"
		   ' response.Write(q)
		   'response.Flush()
		   connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & PRODUCTIONDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), PRODUCTIONDATE)
		If Cek_IsDailyClosing = "FALSE" then
		  q	=	" Insert into PriceListTBS (" & _
			" PRODUCTCODE" & _
			" , ONDATE" & _
			" , PLANTINGYEAR " & _
			" , PRICE" & _
			" , FCENTRY" & _
			" , FCEDIT" & _
			" , FCIP" & _
			" , FCBA" & _
			" , LASTUPDATE" & _
			" , LASTTIME" & _
			" ) values (" & _
			"  '" & PRODUCTCODE & "'" & _
			" , to_date('" & PRODUCTIONDATE & "','mm/dd/yyyy')" & _
			" ," & PLANTINGYEAR & " " & _
			" , '" & To_Nol(PRICE) & "'" & _
			" , '" & Session("CodeName") & "'" & _
			" , '" & Session("CodeName") & "'" & _
			" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
			" , '" & Session("fcBA") & "'" & _
			" , sysdate " & _
			" , To_Char(sysdate, 'HH24:MI')" & _
			" )"
			'response.Write(q)
			'Response.Flush()
			connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & PRODUCTIONDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
        cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), PRODUCTIONDATE)
		If Cek_IsDailyClosing = "FALSE" then
		   q=" Update PriceListTBS set" & _
			 " PRICE = '" & To_Nol(PRICE) & "'" & _
			 " , PLANTINGYEAR = "  & PLANTINGYEAR & " " & _
			 " where" & _
		     " PRODUCTCODE = '" & PRODUCTCODE & "'" & _
			 " And ONDATE = to_date('" & ONDATE & "','mm/dd/yyyy')" & _
			 " and fcBA = '" & Session("fcBA") & "'"
		   connStr.execute(q)
			'response.Write(q)
			'response.Flush()
			connStr.execute("commit")
			'RESPONSE.Write "SQLnsddsdya : " & Q
		else
			err.number = 1001
			err.description = "Tanggal """ & PRODUCTIONDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if len(trim(Perioddate)) <> 0 then
   Perioddate = Request("Perioddate")
else
   Perioddate = textDate(date)
end if

if len(trim(PRODUCTIONDATE)) <> 0 then
   PRODUCTIONDATE = Request("PRODUCTIONDATE")
else
   PRODUCTIONDATE = textDate(date)
end if

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")
If Err.Number <> 0 then
   response.Write  "<div class = ""error"">" & err.description & "</div>"
End if

 '-------------ISI Payment Form
 %>

 <form name="frmEntry" method="post" action="TBSPriceList.asp">
	<table class="data" width="100%" cellpadding="0" cellspacing="0">
          <TR> 
            <td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE TBS PRICE LIST</strong></font></div></td>
          </TR>

		<tr><td align="left">Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="Perioddate" size="10" value="<%=Perioddate%>">
		
		</td></tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Product Code</font></td>
			<td align="left">
				<select name="ProductCODE">
				<%
					Dim rsProductCode, strProductCode
					strProductCode = " SELECT PRODUCTCODE, PRODUCTDESCRIPTION FROM PRODUCT WHERE FCBA = '" &  Session("fcBA") & "' AND PRODUCTCODE = 'FFB' ORDER BY PRODUCTCODE"
					
					Set rsProductCode = Server.CreateObject("ADODB.Recordset")
					rsProductCode.Open strProductCode, connstr
					
					If Not rsProductCode.EOF then
						Do While Not rsProductCode.EOF
				%>
					<option value="<%=rsProductCode("PRODUCTCODE")%>" <%If PRODUCTCODE = rsProductCode("PRODUCTCODE") Then%> selected <%End If%>><%=rsProductCode("PRODUCTCODE")%> : <%=rsProductCode("PRODUCTDESCRIPTION")%></option> 
				<%			rsProductCode.movenext
						Loop
					End if
					rsProductCode.Close
					set rsProductCode = Nothing
				%>
				</select><font color="#FF0000"> *</font>
			</td>
		</tr>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		<a href="p_TBSPriceList.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&Perioddate=<%=Perioddate%>&txt_SHIFTCODE=<%=REQUEST("txt_SHIFTCODE")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		<!-- HADIR -->
		<%
		dim dataModule,dataSubmodule,dataDetail
		dataModule 		= "Harvesting"
		dataSubmodule 	= "PriceListTBS"
		dataDetail 		= "NULL"
		%>
		<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
		<!-- END:HADIR -->
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   
	  q	=	" select" & _
			" ProductCODE, to_char(ONDATE,'mm/dd/yyyy') as ONDATE,  PLANTINGYEAR " & _
			" , PRICE, FCENTRY, FCEDIT, FCIP, FCBA from PriceListTBS where fcBA='" & session("fcBA") & "'" 
			
			
	 
	if  len(trim(Perioddate)) > 0 then
		q = q & " and to_char(ONDATE,'MMYYYY') = to_char(to_date('" & PRODUCTIONDATE & "', 'MM/DD/YYYY'),'MMYYYY')" 
	end if
	if  len(trim(PRODUCTCODE)) > 0 then
		q = q & " and PRODUCTCODE='" & PRODUCTCODE & "'" 
	end if
	q = q  & " ORDER BY PRODUCTCODE, PLANTINGYEAR, ONDATE" 
  ' response.Write(q)
  'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>

   <form action="TBSPriceList.asp" method="post" name="eForm">
   <table class="data" width="100%" cellpadding="0" cellspacing="0">
		<td align="left">Date</td>
		<%IF txtJenis="Edit" Then%>
				<td align="left"><input type="text" maxlength="10" name="PRODUCTIONDATE" size="10" value="<%=PRODUCTIONDATE%>" readonly="True" style="background:#999999">
		<%Else%>
				<td align="left"><input class="datepicker" type="text" maxlength="10" name="PRODUCTIONDATE" size="10" value="<%=PRODUCTIONDATE%>">
		<%End if%>
				<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Product Code</font></td>
			<td align="left">
				<select name="PRODUCTCODE">
				<%
					'rsProductCode, strProductCode
					strProductCode = " SELECT PRODUCTCODE, PRODUCTDESCRIPTION FROM PRODUCT WHERE FCBA = 'PSG' AND PRODUCTCODE = 'FFB' ORDER BY PRODUCTCODE"
					
					Set rsProductCode = Server.CreateObject("ADODB.Recordset")
					rsProductCode.Open strProductCode, connstr
					
					If Not rsProductCode.EOF then
						Do While Not rsProductCode.EOF
				%>
					<option value="<%=rsProductCode("PRODUCTCODE")%>" <%If PRODUCTCODE = rsProductCode("PRODUCTCODE") Then%> selected <%End If%>><%=rsProductCode("PRODUCTCODE")%> : <%=rsProductCode("PRODUCTDESCRIPTION")%></option> 
				<%			rsProductCode.movenext
						Loop
					End if
					rsProductCode.Close
					set rsProductCode = Nothing
				%>
				</select><font color="#FF0000"> *</font>
			</td>
		</tr>
		<tr>
		    <td align="left"><font face="Verdana" size="1">Planting Year</font></td>
			<td align="left">
			    <%
					        Dim rsPlantingDate,strsql
							Set rsPlantingDate = Server.CreateObject("ADODB.Recordset")
							strsql =  " SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM Budget_Field WHERE FCBA = '" & Session("fcBA") & "'"
							'response.write  strsql
							rsPlantingDate.Open strsql, connStr
				%>
				<select name="PLANTINGYEAR">	
						<option value=""></option>
						<%
						
							
							if not rsPlantingDate.eof then
								DO while not rsPlantingDate.eof
								if PLANTINGYEAR = rsPlantingDate("PlantingDate") then
						%>			<option value="<%=rsPlantingDate("PlantingDate")%>" Selected><%=rsPlantingDate("PlantingDate")%></option>
								<%else%>
									<option value="<%=rsPlantingDate("PlantingDate")%>"><%=rsPlantingDate("PlantingDate")%></option>
								<%end if%>
						<%			rsPlantingDate.movenext
								loop
							End if
							rsPlantingDate.Close
							set rsPlantingDate = nothing
						%>
				</select>		
			</td>
		</tr>
		<tr><td align="left">Price</td><td align="left"><input type="text" size="15" name="PRICE" value="<%=PRICE%>"></td></tr>
	    <TR><td colspan="2" align="left">
		  <input type="hidden" name="Perioddate" value="<%=Perioddate%>">
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
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>
   </form>
<!--#include file="validation_show.Asp"-->    
<%
   JumRec=1
   intRecordsPerPage = 20
	
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
   <table width="737" class="data" >
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">PRODUCT CODE</TD>	
	<TD align="left">DATE</TD>	
	<TD align="left">PLANTING YEAR</TD>
	<TD align="left">PRICE</TD>	
	</tr>
   
        <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-5
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td width="79" align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="21">

			<a href="TBSPriceList.asp?Perioddate=<%=rsGAD("ONDATE")%>&ONDATE=<%=rsGAD("ONDATE")%>&PRODUCTCODE=<%=rsGAD("PRODUCTCODE")%>&txttable=<%=request("txttable")%>&PLANTINGYEAR=<%=rsGAD("PLANTINGYEAR")%>&PRICE=<%=rsGAD("PRICE")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="TBSPriceList.asp?Perioddate=<%=rsGAD("ONDATE")%>&ONDATE=<%=rsGAD("ONDATE")%>&PRODUCTCODE=<%=rsGAD("PRODUCTCODE")%>&txttable=<%=request("txttable")%>&PLANTINGYEAR=<%=rsGAD("PLANTINGYEAR")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
		    <%for i = 0 to nCountofField%>
			    <td width="597" align="left" valign="top" title="<%=rsGAD(i).Name%>"><%=rsGAD(i)%></td>
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
