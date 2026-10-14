<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Product Storage Reading</title>
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

<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-docdetail.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-docdetail.css?v=20261009b" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014b"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-doc"<% end if %>>
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
	'Getting Data (FUNCTION)
Function Get_TinggiMaxTerukur(initialStorageCode, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select nvl(maxhighleveltank, 0) maxhighleveltank" & _
	                " from productstorage" & _
					" where storagecode = '" & initialStorageCode & "'" & _
					"       and fcba = '" & initialFcba & "'"
							
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("maxhighleveltank"))
	Else
		initialAvailable = 0
	End If
	
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	Get_TinggiMaxTerukur = initialAvailable
End Function

Function Get_ProductCode(initialStorageCode, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select nvl(upper(product_productcode), 'N/A') product_productcode" & _
	                " from productstorage" & _
					" where storagecode = '" & initialStorageCode & "'" & _
					"       and fcba = '" & initialFcba & "'"
							
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = rsAvailable("product_productcode")
	Else
		initialAvailable = 0
	End If
	
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	Get_ProductCode = initialAvailable
End Function

Function Get_VolumeMaxTerukur(initialStorageCode, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select nvl(maxcapacity, 0) maxcapacity" & _
	                " from productstorage" & _
					" where storagecode = '" & initialStorageCode & "'" & _
					"       and fcba = '" & initialFcba & "'"
							
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("maxcapacity"))
	Else
		initialAvailable = 0
	End If
	
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	Get_VolumeMaxTerukur = initialAvailable
End Function

Function Get_Buttom(initialStorageCode, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select nvl(deadstock_capacity, 0) deadstock_capacity" & _
	                " from productstorage" & _
					" where storagecode = '" & initialStorageCode & "'" & _
					"       and fcba = '" & initialFcba & "'"
							
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("deadstock_capacity"))
	Else
		initialAvailable = 0
	End If
	
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	Get_Buttom = initialAvailable
End Function

Function Get_Konversi(initialStorageCode, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select nvl(weightconvertion, 0) weightconvertion" & _
	                " from productstorage" & _
					" where storagecode = '" & initialStorageCode & "'" & _
					"       and fcba = '" & initialFcba & "'"
							
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("weightconvertion"))
	Else
		initialAvailable = 0
	End If
	
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	Get_Konversi = initialAvailable
End Function

Function Get_SpecificGravity(initialTemperature)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select nvl(specificgravity, 0) specificgravity" & _
	                " from specificgravity" & _
					" where temperature = '" & initialTemperature & "'"
							
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("specificgravity"))
	Else
		initialAvailable = 0
	End If
	
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	Get_SpecificGravity = initialAvailable
End Function

Function Get_Round(initialNumber, initialLenght)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select nvl(round('" & initialNumber & "', '" & initialLenght & "'), 0) valueRounding" & _
	                " from dual"
							
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("valueRounding"))
	Else
		initialAvailable = 0
	End If
	
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	Get_Round = initialAvailable
End Function

Function Get_IsReadytoDeliver(initialStorageCode, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select nvl(upper(isreadytodeliver), 'N') isreadytodeliver" & _
	                " from productstorage" & _
					" where storagecode = '" & initialStorageCode & "'" & _
					"       and fcba = '" & initialFcba & "'"
							
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = rsAvailable("isreadytodeliver")
	Else
		initialAvailable = 0
	End If
	
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	Get_IsReadytoDeliver = initialAvailable
End Function
%>


<%
Dim FCENTRY, FCEDIT, FCIP, FCBA
Dim READINGDATE, PSTORAGE_STORAGECODE, QUANTITY,FFA_QUALITY,MOISTURE_QUALITY,DIRT_QUALITY,DOCNUMBER,NORMALSOUNDINGNUMBER
dim txt_STORAGECODE, Perioddate,not_inserting

Dim cmdView,rsPR

Dim oRS, oSQL
Dim tempProductofStorage : tempProductofStorage = ""
Dim txtTinggi : txtTinggi = Request("txtTinggi")
Dim txtSuhu : txtSuhu = Request("txtSuhu")
Dim TinggiMaxTerukur : TinggiMaxTerukur = 0
Dim VolumeMaxTerukur : VolumeMaxTerukur = 0
Dim Buttom : Buttom = 0
Dim Konversi : Konversi = 0
Dim SpecificGravity : SpecificGravity = 0
Dim tempQuantityKg : tempQuantityKg = 0
Dim IsReadytoDeliver : IsReadytoDeliver = ""
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

READINGDATE = Request("READINGDATE")
PSTORAGE_STORAGECODE = Request("PSTORAGE_STORAGECODE")
QUANTITY = Request("QUANTITY")
FFA_QUALITY = Request("FFA_QUALITY")
MOISTURE_QUALITY = Request("MOISTURE_QUALITY")
DIRT_QUALITY = Request("DIRT_QUALITY")
DOCNUMBER = Request("DOCNUMBER")
txttable = Request("txttable")
NORMALSOUNDINGNUMBER = Request("NORMALSOUNDINGNUMBER")

	dim read_only,hidden_only
	if len(trim(DOCNUMBER)) = 0 AND len(trim(NORMALSOUNDINGNUMBER)) = 0 then 
		hidden_only = "style='display:none'"
		read_only = ""
	else
		hidden_only = ""
		read_only = "readonly='True' style='background:#FFFFCC'"
	end if
	
if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then	
	if Request("txtJenis")= "Delete" then
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), READINGDATE)
		If Cek_IsDailyClosing = "FALSE" then
			q = " delete from productstoragereading" & _
				" where readingdate = to_date('" & READINGDATE & "','mm/dd/yyyy')" & _
				"       and pstorage_storagecode='" & PSTORAGE_STORAGECODE & "'" & _
				"       and fcba = '" & Session("fcBA") & "'"
		   'response.Write(q)
		   connStr.execute(q)
		   
			if err.number = 0 then
			   connStr.CommitTrans
			end if
		else
			err.number = 1001
			err.description = "Tanggal """ & READINGDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
		cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), READINGDATE)
		If Cek_IsDailyClosing = "FALSE" then
			'Definisi Produk untuk Storage yang di maksud
			tempProductofStorage = Get_ProductCode(PSTORAGE_STORAGECODE, Session("fcBA"))
			TinggiMaxTerukur = Get_TinggiMaxTerukur(PSTORAGE_STORAGECODE, Session("fcBA"))
			VolumeMaxTerukur = Get_VolumeMaxTerukur(PSTORAGE_STORAGECODE, Session("fcBA"))
			Buttom = Get_Buttom(PSTORAGE_STORAGECODE, Session("fcBA"))
			Konversi = Get_Konversi(PSTORAGE_STORAGECODE, Session("fcBA"))
			SpecificGravity = Get_SpecificGravity(txtSuhu)
			IsReadytoDeliver = Get_IsReadytoDeliver(PSTORAGE_STORAGECODE, Session("fcBA"))
			
			'Response.Write TinggiMaxTerukur & "<br/>" & txtTinggi
			
			if len(trim(DOCNUMBER)) <= 0 then 	
					
					'response.write(QUANTITY)
					
						q = " insert into productstoragereading (" & _
							" readingdate, pstorage_storagecode, quantity, fcentry, fcedit" & _
							" , fcip, fcba, lastupdate, lasttime, ffa_quality" & _
							" , moisture_quality, dirt_quality, isautosounding, highleveltank, temperature" & _
							" ) values (" & _
							" to_date('" & READINGDATE & "', 'mm/dd/yyyy')" & _
							" , '" & PSTORAGE_STORAGECODE & "'" & _
							" , '" & QUANTITY & "'" & _
							" , '" & Session("CodeName") & "'" & _
							" , '" & Session("CodeName") & "'" & _
							" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
							" , '" & Session("fcBA") & "'" & _
							" , sysdate" & _
							" , to_char(sysdate,'hh24:mi')" & _
							" , '" & FFA_QUALITY & "'" & _
							" , '" & MOISTURE_QUALITY & "'" & _
							" , '" & DIRT_QUALITY & "'" & _
							" , 'FALSE'" & _
							" , '" & txtTinggi & "'" & _
							" , '" & txtSuhu & "'" & _
							" )"
						'response.Write(q)
						'response.Flush()
						connStr.execute(q)
				
			ELSE	
			
				q = " insert into productstoragereading (" & _
							" readingdate, pstorage_storagecode, quantity, fcentry, fcedit" & _
							" , fcip, fcba, lastupdate, lasttime, ffa_quality" & _
							" , moisture_quality, dirt_quality, isautosounding, highleveltank, temperature, DOCNUMBER" & _
							" ) values (" & _
							" to_date('" & READINGDATE & "', 'mm/dd/yyyy')" & _
							" , '" & PSTORAGE_STORAGECODE & "'" & _
							" , '" & QUANTITY & "'" & _
							" , '" & Session("CodeName") & "'" & _
							" , '" & Session("CodeName") & "'" & _
							" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
							" , '" & Session("fcBA") & "'" & _
							" , sysdate" & _
							" , to_char(sysdate,'hh24:mi')" & _
							" , '" & FFA_QUALITY & "'" & _
							" , '" & MOISTURE_QUALITY & "'" & _
							" , '" & DIRT_QUALITY & "'" & _
							" , 'FALSE'" & _
							" , '" & txtTinggi & "'" & _
							" , '" & txtSuhu & "'" & _
							" , '" & DOCNUMBER & "'" & _
							" )"
						'response.Write(q)
						'response.Flush()
						connStr.execute(q)

			End If	
			if err.number = 0 then
			   connStr.CommitTrans
			end if
		
		else
			err.number = 1001
			err.description = "Tanggal """ & READINGDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
		cmdView	= "True"
		
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), READINGDATE)
		If Cek_IsDailyClosing = "FALSE" then
			'Definisi Produk untuk Storage yang di maksud
			tempProductofStorage = Get_ProductCode(PSTORAGE_STORAGECODE, Session("fcBA"))
			TinggiMaxTerukur = Get_TinggiMaxTerukur(PSTORAGE_STORAGECODE, Session("fcBA"))
			VolumeMaxTerukur = Get_VolumeMaxTerukur(PSTORAGE_STORAGECODE, Session("fcBA"))
			Buttom = Get_Buttom(PSTORAGE_STORAGECODE, Session("fcBA"))
			Konversi = Get_Konversi(PSTORAGE_STORAGECODE, Session("fcBA"))
			SpecificGravity = Get_SpecificGravity(txtSuhu)
			IsReadytoDeliver = Get_IsReadytoDeliver(PSTORAGE_STORAGECODE, Session("fcBA"))
			
			'Response.Write TinggiMaxTerukur & "<br/>" & txtTinggi
			if len(trim(DOCNUMBER)) <= 0 then 	

						q = " update productstoragereading" & _
							" set quantity = '" & QUANTITY & "'" & _
							"     , ffa_quality = '" & FFA_QUALITY & "'" & _
							"     , moisture_quality = '" & MOISTURE_QUALITY & "'" & _
							"     , dirt_quality = '" & DIRT_QUALITY & "'" & _
							"     , isautosounding = 'FALSE'" & _
							"     , lastupdate = sysdate " & _
							"     , lasttime = to_char(sysdate,'hh24:mi')" & _
							"     , highleveltank = '" & txtTinggi & "'" & _
							"     , temperature = '" & txtSuhu & "'" & _
							" where readingdate = to_date('" & READINGDATE & "', 'mm/dd/yyyy')" & _ 
							"       and pstorage_storagecode = '" & PSTORAGE_STORAGECODE & "'" & _
							"       and fcba = '" & Session("fcBA") & "'"
							
						'response.Write(q)
						connStr.execute(q)
						
						'response.Flush()
						connStr.execute("commit")
						'RESPONSE.Write "SQLnsddsdya : " & Q
				
			ELSE
			
			
			
			
			
			End If
			if err.number = 0 then
			   connStr.CommitTrans
			end if
		else
			err.number = 1001
			err.description = "Tanggal """ & READINGDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
		cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if

Dim txtModule, txtTable, txtSearch, txtField, i, q, nCountofField, NO, txtJenis

if len(trim(READINGDATE)) <> 0 then
   READINGDATE = Request("READINGDATE")
else
   READINGDATE = textDate(date)
end if

txtTable  = Request("txtTable")
txtModule = Request("txtModule")
txtSearch = Request("txtSearch")
txtField  = Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")

If Err.Number <> 0 then
	connStr.RollbackTrans
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if
%>
<script language="javascript">
			function level_tank(str)
				{
				var tempr = document.getElementById("txtSuhu").value;
				if 	(tempr == ""){
					document.getElementById("QUANTITY").value="Temperature ?";
					}else{
						if (str.length==0)
						  {
						  document.getElementById("QUANTITY").value="0";
						  return;
						  }
						  
						if (window.XMLHttpRequest)
						  {// code for IE7+, Firefox, Chrome, Opera, Safari
						  xmlhttp=new XMLHttpRequest();
						  }
						else
						  {// code for IE6, IE5
						  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
						  }
						  
						xmlhttp.onreadystatechange=function()
						  {
						  if (xmlhttp.readyState==4 && xmlhttp.status==200)
							{
							document.getElementById("QUANTITY").value=xmlhttp.responseText;
							}
						  }
						xmlhttp.open("GET","ajax/sounding_ajx.asp?initialStorageCode=<%=PSTORAGE_STORAGECODE%>&initialReadingNumber="+str+"&initialFcba=<%=session("fcBA")%>&initialTemperature=" + tempr,true);
						xmlhttp.send();
					}
				}
				function level_tempr(str)
				{
				var hlt = document.getElementById("txtTinggi").value;
				if 	(hlt == ""){
					document.getElementById("QUANTITY").value="Reading Lvl Tank ?";
					}else{
						if (str.length==0)
						  {
						  document.getElementById("QUANTITY").value="0";
						  return;
						  }
						  
						if (window.XMLHttpRequest)
						  {// code for IE7+, Firefox, Chrome, Opera, Safari
						  xmlhttp=new XMLHttpRequest();
						  }
						else
						  {// code for IE6, IE5
						  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
						  }
						  
						xmlhttp.onreadystatechange=function()
						  {
						  if (xmlhttp.readyState==4 && xmlhttp.status==200)
							{
							document.getElementById("QUANTITY").value=xmlhttp.responseText;
							}
						  }
						xmlhttp.open("GET","ajax/sounding_ajx.asp?initialStorageCode=<%=PSTORAGE_STORAGECODE%>&initialReadingNumber="+hlt+"&initialFcba=<%=session("fcBA")%>&initialTemperature=" + str,true);
						xmlhttp.send();
					}
				}
</script>
<!-------------ISI Payment Form!------------->
	<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-tachometer" aria-hidden="true"></i>Header Product Storage Reading</div><% end if %>
	<form name="frmEntry" method="post" action="PRODUCTSTORAGEREADING.asp">
		<table class="data" width="70%" cellpadding="0" cellspacing="0">
			<tr> 
				<td colspan="2">
					<div align="left">
						<font size="2" face="Verdana"><strong>MAINTENANCE STORAGE READING</strong></font>
					</div>
				</td>
			</tr>
			<tr>
				<td align="left">Date</td>
				<td align="left">
					<input class="datepicker" type="text" maxlength="10" name="READINGDATE" size="10" value="<%=READINGDATE%>">
				</td>
			</tr>
			<tr>
				<td align="left">Storage Code</td>
				<td align="left">
					<select name="PSTORAGE_STORAGECODE">
						<%cmdCallComboSetup "PRODUCTSTORAGE",REQUEST("PSTORAGE_STORAGECODE"),"fcba",session("fcba")%>
					</select>
				</td>
			</tr>
			<tr>
				<td colspan="2"  align="left" class="fbp-doc-actions">
					<input Type="Hidden" Value="True" Name="cmdView">
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
					<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
					<a href="PrintPRODUCTSTORAGEREADING.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&Perioddate=<%=Perioddate%>&txt_STORAGECODE=<%=REQUEST("txt_STORAGECODE")%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
					<%
					dim dataModule,dataSubmodule,dataDetail
					dataModule 		= "Mill Production"
					dataSubmodule 	= "Storage Reading"
					dataDetail 		= "NULL"
					%>
					<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
					<!-- END:HADIR -->
				</td>
			</tr>
		</table>  
	</form>
<%
	if cmdView = "True" and Len(Trim(READINGDATE)) > 0 and Len(Trim(PSTORAGE_STORAGECODE)) > 0 Then
	Dim rsGAD,intRecordsPerPage,cColor
	set rsGAD = server.CreateObject("Adodb.recordset")
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q = " select readingdate" & _
	    "        , pstorage_storagecode" & _
	    "        , quantity" & _
	    "        , ffa_quality" & _
	    "        , moisture_quality" & _
	    "        , dirt_quality" & _
		"        , highleveltank" & _
		"        , temperature" & _
		"        , DOCNUMBER" & _
		"        , NORMALSOUNDINGNUMBER" & _
	    " from productstoragereading" & _
	    " where fcBA = '" & session("fcBA") & "'" & _
	    "       and nvl(isautosounding,'FALSE') = 'FALSE'" & _
	    "       and READINGDATE = to_date('" & READINGDATE & "', 'MM/DD/YYYY')" & _
	    "       and PSTORAGE_STORAGECODE = '" & PSTORAGE_STORAGECODE & "'"
	'response.Write(q)
	'response.Flush()
	rsGAD.CursorType = 3
	rsGAD.CursorLocation = 3
	rsGAD.Open q, connstr

	tempProductofStorage = Get_ProductCode(PSTORAGE_STORAGECODE, Session("fcBA"))
	IsReadytoDeliver = Get_IsReadytoDeliver(PSTORAGE_STORAGECODE, Session("fcBA"))
	if not rsGAD.EOF then
		IF request("txtJenis") <> "Edit" then
			not_inserting = "display:none;"
		else
			not_inserting = ""
		end if
	end if
	
%>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="PRODUCTSTORAGEREADING.asp" method="post" name="eForm" style="<%=not_inserting%>">
   <table class="data" cellpadding="0" cellspacing="0" width="70%" >
   <tr <%=hidden_only%>>
			<td align="left">Normal Sounding Number</td>
			<td align="left">
				<input type="text" size="20" name="NORMALSOUNDINGNUMBER" value="<%=NORMALSOUNDINGNUMBER%>" <%=read_only%>>
			</td>
		</tr>
		<tr <%=hidden_only%>>
			<td align="left">Document Number</td>
			<td align="left">
				<input type="text" size="20" name="DOCNUMBER" value="<%=DOCNUMBER%>" <%=read_only%>>
			</td>
		</tr>
		<tr <%=hidden_only%>>
			<td align="left">Reading Level Tank</td>
			<td align="left">
				<input id="txtTinggi" onkeyup="level_tank(this.value)" type="text" size="20" name="txtTinggi" value="<%=txtTinggi%>" <%=read_only%>>
			</td>
		</tr>
		
		<tr <%=hidden_only%>>
			<td align="left">Temperature</td>
			<td align="left">
				<input id="txtSuhu" onkeyup="level_tempr(this.value)" type="text" size="20" name="txtSuhu" value="<%=txtSuhu%>" <%=read_only%>>
			</td>
		</tr>
		<tr>
			<td align="left">Quantity (KG)</td>
			<td align="left">
				<input id="QUANTITY" type="text" size="20" name="QUANTITY" value="<%=QUANTITY%>" <%=read_only%>>
			</td>
		</tr>
		<tr>
			<td align="left">FFA (%)</td>
			<td align="left">
				<input type="text" size="20" name="FFA_QUALITY" value="<%=FFA_QUALITY%>" <%=read_only%>>
			</td>
		</tr>
		<tr>
			<td align="left">MOISTURE (%)</td>
			<td align="left">
				<input type="text" size="20" name="MOISTURE_QUALITY" value="<%=MOISTURE_QUALITY%>" <%=read_only%>>
			</td>
		</tr>
		<tr>
			<td align="left">DIRT (%)</td>
			<td align="left">
				<input type="text" size="20" name="DIRT_QUALITY" value="<%=DIRT_QUALITY%>" <%=read_only%>>
			</td>
		</tr>
	    <TR>
			<td colspan="2" align="left">
				<input type="hidden" name="READINGDATE" value="<%=READINGDATE%>">
				<input type="hidden" name="PSTORAGE_STORAGECODE" value="<%=PSTORAGE_STORAGECODE%>">
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
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Data</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines" >
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">DATE</TD>
	<TD align="left">STORAGE CODE</TD>
	<TD align="left">QUANTITY (KG)</TD>
	<TD align="left">FFA (%)</TD>
	<TD align="left">MOISTURE (%)</TD>
	<TD align="left">DIRT (%)</TD>
	<TD align="left">READING LEVEL TANK</TD>
	<TD align="left">TEMPERATURE</TD>
	<TD align="left">DOC NUMBER</TD>
	<TD align="left">NORMAL SOUNDING</TD>
	</tr>
	</thead>
	<tbody>
        <%
   if Not rsGAD.eof then
      nCountofField = rsGAD.fields.count - 1
	  no = 1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td width="79" align="left" valign="top"><%=NO%></td>																																					
			<td valign="middle" width="21">
			<a href="PRODUCTSTORAGEREADING.asp?READINGDATE=<%=rsGAD("READINGDATE")%>&PSTORAGE_STORAGECODE=<%=rsGAD("PSTORAGE_STORAGECODE")%>&QUANTITY=<%=rsGAD("QUANTITY")%>&FFA_QUALITY=<%=rsGAD("FFA_QUALITY")%>&MOISTURE_QUALITY=<%=rsGAD("MOISTURE_QUALITY")%>&DIRT_QUALITY=<%=rsGAD("DIRT_QUALITY")%>&txtTinggi=<%=rsGAD("HIGHLEVELTANK")%>&txtSuhu=<%=rsGAD("TEMPERATURE")%>&DOCNUMBER=<%=rsGAD("DOCNUMBER")%>&NORMALSOUNDINGNUMBER=<%=rsGAD("NORMALSOUNDINGNUMBER")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="PRODUCTSTORAGEREADING.asp?READINGDATE=<%=rsGAD("READINGDATE")%>&PSTORAGE_STORAGECODE=<%=rsGAD("PSTORAGE_STORAGECODE")%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0"></a>
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
   </tbody>
   </TABLE>
   </div>
<%End if%>  
 
  <!--------------akhir!-------------->
</div></div>
  <!--#include file="bottom.jsp"-->
</body>
</html>
