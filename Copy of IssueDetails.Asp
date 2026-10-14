<%Option Explicit%>
<%On Error Resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<html>
<head>
<title>iPS - Stock Issue</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src="date-picker.js"></script> 
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
<table width="100%">
 <tr>
 <td width="22%" valign="top">
 <!--#include file="menu.asp"-->
 </td>
 <td width="78%" valign="top">
 <%
 '-------------awal-----------------------
 Dim xCode,fcCode,cmdView,PoCode,IssueDate,ITEMCODE,QUANTITY,AMOUNT, LineNoHist
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
 Dim Jobcode, LocationType, LocationCode, Remarks, Reference
 Dim qLineMax, rsqLineMax, tempMaxLine
 Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
 Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand, QuantityOriginal
 Dim done, qq  
 Dim qLessZero, rsLessZero, tempislesszero
 Dim qLineUpper, rsqLineUpper, totalUpper
 Dim qQuantity, rsqQuantity, totalQuantity,LineNoKey
 
 Dim rsReceiveDetail, txtStoreCode
 Dim v_CrossCekDetail, txtLineNoHist, originalQuantity
 Dim txtItemCodeDelete
 Dim qDeleteHist
 Dim strLineMax, rsstrLineMax, tempMaxLineNo , txtLineNoKey


 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 if request("fccOde")="" then
    fcCode="" '----- Nantinya Ambil dari Format Master
 else
    fcCode = request("fccOde")
 end if
 
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if Request("IssueDate")<>"" then
   IssueDate   =Request("IssueDate")
 else
   IssueDate   =textDate(date)
 end if
 xCode = Request("xCode")
 
 '---- Data Details
 JobCode=Request("JobCode")
 LocationType=Request("LocationType")
 LocationCode=Request("LocationCode")
 Remarks = Request("Remarks")
 Reference = Request("Reference")
 ITEMCODE=Request("ITEMCODE")
 LineNoKey=Request("LineNoKey")
 QUANTITY=Request("QUANTITY")
 AMOUNT =Request("AMOUNT")
 Linenohist = Request ("linenohist")
 
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Header" then
	   connstr.BeginTrans
       If err.number=0 then
	      q="Delete from Issue where  SivCODE='" & fcCode & "' and fcBA='" & Session("fcBA") & "'"
   	      connStr.execute(q)
		  if err.number  = 0 then
		     connStr.CommitTrans
		  end if
	   end if
       cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
	
	    connstr.BeginTrans
	   
		q = "Select storecode from Issue where SivCODE='" & fcCode & _
			"' and fcBA='" & Session("fcBA") & "'"
			
			
		Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
		rsReceiveDetail.Open q, connstr
				
		IF Not rsReceiveDetail.eof then
			txtStoreCode = rsReceiveDetail("STORECODE")
		End if
		rsReceiveDetail.Close
		set rsReceiveDetail = Nothing		   
	   
	   
		q = "Select stockcode, linenohist, quantity, lineNoKey from IssueDetails where SivCODE='" & fcCode & _
			"' and LineNoKey='" & request("LineNoKey") & _
			"' and STOCKCODE='" & request("ItemCode") & _
			"' and fcBA='" & Session("fcBA") & "' " & _
			"  and linenohist = " & Linenohist
		

		Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
		rsReceiveDetail.Open q, connstr
		
	
				
		IF Not rsReceiveDetail.eof then
			txtLineNoHist = rsReceiveDetail("LINENOHIST")
			originalQuantity = rsReceiveDetail("QUANTITY")
			txtItemCodeDelete = rsReceiveDetail("stockcode")
			txtLineNoKey = rsReceiveDetail("LineNoKey")
		End if
		rsReceiveDetail.Close
		set rsReceiveDetail = Nothing			

		v_CrossCekDetail = cmdCheckHist(txtLineNoHist, Request("ItemCode"))
	
		if v_CrossCekDetail = 0 then 					

			q =  "Delete from stocktranshist " & _
				 "where " & _
				 "lineno = " & txtLineNoHist & " and " & _
				 "fcba = '" & Session("fcBA") & "'"
		
			connStr.execute(q)
			
			
			
			q = " update stockonhand " & _
				" set " & _
				" quantity = quantity + " & originalQuantity & " " & _
				"where " & _
				"storecode = '" & txtStoreCode & "' and " & _
				"stockcode  = '" & txtItemCodeDelete & "' and " & _ 
				"fcba = '" & session("fcBA") &"' "

				
			
			connStr.execute(q)		
			
								
			q="Delete from IssueDetails where  SivCODE='" & fcCode & "' and LineNoKey=" & LineNoKey & " and fcBA='" & Session("fcBA") & "'" & _
			   "  and linenohist = " & Linenohist
			connStr.execute(q)
			if err.number  = 0 then
			   connStr.CommitTrans
			end if
		else
			done = 2
		end if

		cmdView	= "True"
	elseif Request("txtJenis")="New" then
	  Dim rsPO
	  Set rsPO = Server.CreateObject("ADODB.Recordset")
	  rsPO.Open "Select sivcode from Issue where SIVCODE='" & fcCode & "'", connstr
	  connstr.BeginTrans
	  IF rsPO.eof then	    
	  
	     q="Insert into Issue (SIVCODE, STORECODE, SIVDATE, FCENTRY, FCEDIT, FCIP, FCBA) " & _
		 "values ('" & fcCode & "','" &  xCode & "', to_date('" & IssueDate & "','mm/dd/yyyy') " & _
		 ",'" &  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		 Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
	     connStr.execute(q)
		 'response.Write(q)
	  End if
	  rsPO.Close
	  if err.number=0 then	
	    if err.number  = 0 then
		
			' Add by NAF to handle Inventory Stock On Hand
			' Insert the History of Received Inventory Transaction
			
			qLessZero = "select islesszero from stock " & _	   
				"where " & _
				"fccode = '" & ITEMCODE & "' and " & _
				"fcba = '" & session("fcBA") & "'" 
			
			set rsLessZero = Server.CreateObject("ADODB.Recordset")
			rsLessZero.Open qLessZero, connstr
			if not rsLessZero.eof then
				tempislesszero = rsLessZero("islesszero")
			else
				tempislesszero = "FALSE"
			end if 
			
			
			if err.number = 0 then 			
				qTotalRecordOnHand = "select count(*) totalRecordOnHand from stockonhand " & _
									 "where " & _
									 "storecode = '" & xCode & "' and " & _												 
									 "stockcode = '" & ITEMCODE & "' and " & _
									 "fcba = '" &  session("fcBA") & "'"

				set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
				rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
				if not rsTotalRecordOnHand.eof then
					tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
				else
					tempTotalRecordOnHand = 0
				end if 
				
				done = 0
				
				if tempTotalRecordOnHand > 0 or tempislesszero = "TRUE"  then 
					' We have to Update the Quantity by Adding them with the New one.
					
					
					qQuantityOnHand = "select quantity from stockonhand " & _
									  "where " & _
									  "storecode = '" & xCode & "' and " & _
									  "stockcode  = '" & ITEMCODE & "' and " & _ 
									  "fcba = '" & session("fcBA") &"' "
									  
					set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
					rsQuantityOnHand.Open qQuantityOnHand, connstr
					if not rsQuantityOnHand.eof then
						tempQuantityOnHand = cdbl(rsQuantityOnHand("quantity"))
					else
						tempQuantityOnHand = 0
					end if 		

					
					if tempQuantityOnHand - Quantity >= 0 or tempislesszero = "TRUE"  then 
					
						qLineMax = "select nvl(max(lineno),0) tempMaxLine from stocktranshist " & _
						"where " & _
						"fcba = '" & session("fcBA") & "'" 
						
						set rsqLineMax = Server.CreateObject("ADODB.Recordset")
						rsqLineMax.Open qLineMax, connstr
						if not rsqLineMax.eof then
							tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
						else
							tempMaxLine = 0
						end if 

						tempMaxLine = tempMaxLine + 1
							
						
						q = "Insert into StockTransHist ( " & _
							" LINENO, FCBA, STOCKCODE, TRANSDATE, TRANSTYPE, STORECODE, SOURCETYPE, SOURCECODE, TARGETTYPE, TARGETCODE, QUANTITY) values ( "& _
							  tempMaxLine & ",'" & _ 
							  session("fcBA") & "','" & _
							  ITEMCODE &"'," & _
							  "to_date('" & IssueDate & "','mm/dd/yyyy') ,'" & _
							  "ST_ISSUED" & "','" & _
							  xCode & "'," & _
							  "'ST','" & _
							  xCode & "','" & _
							  LOCATIONTYPE & "','" & _
							  LOCATIONCODE & "'," & _
							  -1 * Quantity & " ) "
														
						connStr.execute(q)
						
						
						' Added by NAF, 2012-12-01
						 strLineMax = " select nvl(max(linenokey),0) maxLineNoKey from IssueDetails " & _
								   " where " & _
								   " fcba = '" & Session("fcBA") & "'"
								   
					  
					
						Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
						rsstrLineMax.CursorType = 3
						rsstrLineMax.CursorLocation = 3
						rsstrLineMax.Open strLineMax, connstr
						
						if not rsstrLineMax.eof then
							tempMaxLineNo = cdbl(rsstrLineMax("maxLineNoKey")) + 1
						end if
									
						rsstrLineMax.Close
						set rsstrLineMax = Nothing 
						' End of Added by NAF, 2012-12-01

						
						
						
						
					    q=  "Insert into  IssueDetails (SIVCODE, STOCKCODE, LOCATIONCODE, LOCATIONTYPE, QUANTITY, PRICE, AMOUNT, " & _
	                        "JOBCODE, REFERENCE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, linenohist, LineNoKey ) " & _
							"values ('" & fcCode & "','" & ITEMCODE & "','"  &  LOCATIONCODE & "','" & LOCATIONTYPE & "'" & _
							",'" & Quantity & "',0,0,'" & Jobcode & "','" &  Reference & "','" & Remarks & "','" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',"& tempMaxLine & "," &tempMaxLineNo & " )"
							'response.Write(q)
						
						connStr.execute(q)
						
						q = "update stockonhand " & _
							"set quantity = quantity - " & Quantity & " " & _ 
							"where " & _
							"storecode = '" & xCode & "' and " & _
							"stockcode  = '" & ITEMCODE & "' and " & _ 
							"fcba = '" & session("fcBA") &"' "
							 
						connStr.execute(q)								

				    else
						done = 1		
					end if
				else
				    if tempislesszero = "TRUE" then 
						' Just Insert the New Record with the following/received Quantity
						' Raise the Error 
						qLineMax = "select nvl(max(lineno),0) tempMaxLine from stocktranshist " & _
										"where " & _
										"fcba = '" & session("fcBA") & "'" 
							
						set rsqLineMax = Server.CreateObject("ADODB.Recordset")
						rsqLineMax.Open qLineMax, connstr
						if not rsqLineMax.eof then
							tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
						else
							tempMaxLine = 0
						end if 

						tempMaxLine = tempMaxLine + 1
							
						
						qq = "Insert into StockTransHist ( " & _
							" LINENO, FCBA, STOCKCODE, TRANSDATE, TRANSTYPE, STORECODE, SOURCETYPE, SOURCECODE, TARGETTYPE, TARGETCODE, QUANTITY) values ( "& _
							  tempMaxLine & ",'" & _ 
							  session("fcBA") & "','" & _
							  ITEMCODE &"'," & _
							  "to_date('" & IssueDate & "','mm/dd/yyyy') ,'" & _
							  "ST_ISSUED" & "','" & _
							  xCode & "'," & _
							  "'ST','" & _
							  xCode & "','" & _
							  LOCATIONTYPE & "','" & _
							  LOCATIONCODE & "'," & _
							  -1 * Quantity & " ) "
														
						connStr.execute(qq)
						
						
						' Added by NAF, 2012-12-01
						 strLineMax = " select max(nvl(linenokey,0)) maxLineNoKey from IssueDetails " & _
								   " where " & _
								   " fcba = '" & Session("fcBA") & "'"
								   
					  response.write strLineMax
					
						Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
						rsstrLineMax.CursorType = 3
						rsstrLineMax.CursorLocation = 3
						rsstrLineMax.Open strLineMax, connstr
						
						if not rsstrLineMax.eof then
							tempMaxLineNo = cdbl(rsstrLineMax("maxLineNoKey")) + 1
						end if
									
						rsstrLineMax.Close
						set rsstrLineMax = Nothing 
						' End of Added by NAF, 2012-12-01						
						
						

					    q=  "Insert into  IssueDetails (SIVCODE, STOCKCODE, LOCATIONCODE, LOCATIONTYPE, QUANTITY, PRICE, AMOUNT, " & _
	                        "JOBCODE, Reference, Remarks, FCENTRY, FCEDIT, FCIP, FCBA, linenohist, lineNoKey ) " & _
							"values ('" & fcCode & "','" & ITEMCODE & "','"  &  LOCATIONCODE & "','" & LOCATIONTYPE & "'" & _
							",'" & Quantity & "',0,0,'" & Jobcode & "','" & Reference & "','" & Remarks & "','" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',"& tempMaxLine & "," &tempMaxLineNo & " )"
							response.Write(q)
						
						connStr.execute(q)						
						
						
						
						
						
						' Just Insert the New Record with the following/received Quantity
						q = "Insert into stockonhand ( " & _
							" STORECODE, STOCKCODE, FCBA, QUANTITY ) values ( '" & _
							xCode & "','" & _
							ITEMCODE & "','" & _
							session("fcBA") & "'," & _
							-1 * Quantity & " ) "
							
							
						connStr.execute(q)		
				    else
					    done = 1
					end if
				end if
			end if
		
	    end if
		
        cmdView	= "True"
     end if		
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	    connstr.BeginTrans
		
		
	    totalUpper = 0
		
	    qLineUpper =   "select count(*) totalUpper from stocktranshist " & _
					   "where " & _
					   "lineno > " & linenohist & " and " & _
					   "fcba = '" & Session("fcBA") & "' and " & _
					   "stockcode = '" & ItemCode & "' " 
					   
		'response.write qLineUpper
		set rsqLineUpper = Server.CreateObject("ADODB.Recordset")
		rsqLineUpper.Open qLineUpper, connstr
		
		if not rsqLineUpper.eof then
			totalUpper = cdbl(rsqLineUpper("totalUpper"))
		else
			totalUpper = 0
		end if 

		if totalUpper = 0 then 

			
			qQuantity = "select quantity  from stocktranshist " & _
				"where " & _
				"lineno = " & linenohist & " and " & _ 
				"fcba = '" & Session("fcBA") & "'"
				
			
			set rsqQuantity= Server.CreateObject("ADODB.Recordset")
			rsqQuantity.Open qQuantity, connstr

			if not rsqQuantity.eof then
				totalQuantity = cdbl(rsqQuantity("quantity"))
			else
				totalQuantity = 0
			end if 
			
			' Add by NAF to handle Inventory Stock On Hand
			' Insert the History of Received Inventory Transaction
			

			
			qLessZero = "select islesszero from stock " & _	   
				"where " & _
				"fccode = '" & ITEMCODE & "' and " & _
				"fcba = '" & session("fcBA") & "'" 
			
			
			set rsLessZero = Server.CreateObject("ADODB.Recordset")
			rsLessZero.Open qLessZero, connstr
			if not rsLessZero.eof then
				tempislesszero = rsLessZero("islesszero")
			else
				tempislesszero = "FALSE"
			end if 

			
			
			qQuantityOnHand = "select quantity from stockonhand " & _
							  "where " & _
							  "storecode = '" & xCode & "' and " & _
							  "stockcode  = '" & ITEMCODE & "' and " & _ 
							  "fcba = '" & session("fcBA") &"' "
							  
			response.write 	qQuantityOnHand			  
			set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
			rsQuantityOnHand.Open qQuantityOnHand, connstr
			if not rsQuantityOnHand.eof then
				QuantityOriginal = cdbl(rsQuantityOnHand("quantity"))
			else
				QuantityOriginal = 0
			end if 	
			


			if ((tempislesszero = "TRUE") or ((QuantityOriginal - totalQuantity - Quantity) >= 0)) then 
			
				qDeleteHist = "delete stocktranshist " & _
				" where " & _
				" lineno = " & linenohist & " and " & _ 
				" fcba = '" & Session("fcBA") & "'"
				

				connStr.execute(qDeleteHist)


				qq = "Insert into StockTransHist ( " & _
					" LINENO, FCBA, STOCKCODE, TRANSDATE, TRANSTYPE, STORECODE, SOURCETYPE, SOURCECODE, TARGETTYPE, TARGETCODE, QUANTITY) values ( "& _
					  linenohist & ",'" & _ 
					  session("fcBA") & "','" & _
					  ITEMCODE &"'," & _
					  "to_date('" & IssueDate & "','mm/dd/yyyy') ,'" & _
					  "ST_ISSUED" & "','" & _
					  xCode & "'," & _
					  "'ST','" & _
					  xCode & "','" & _
					  LOCATIONTYPE & "','" & _
					  LOCATIONCODE & "'," & _
					  -1 * Quantity & " ) "
												
				connStr.execute(qq)
		   
		   
		   
				q =  "Update IssueDetails set Quantity= " & Quantity & _ 
					" where  SIVCODE='" & fcCode & "' and JobCode='" & JobCode & "'" & _
					"  and Locationtype='" & Locationtype  & "' and Locationcode='" & Locationcode & "'" & _
					"  and LineNoKey = " & LineNoKey & " and StockCode='" & ItemCode & "' and fcBA='" & Session("fcBA") & "'"
					
				connStr.execute(q)
				
				
				
				q = "update stockonhand " & _
					" set quantity = Quantity - " & totalQuantity & " + " & -1 * Quantity & _ 
					" where " & _
					" storecode = '" & xCode & "' and " & _
					" stockcode  = '" & ITEMCODE & "' and " & _ 
					" fcba = '" & session("fcBA") &"' "
					 
				connStr.execute(q)			
				
				
				if err.number  = 0 then
					connStr.CommitTrans
				end if
				done = 0
			else
			    done = 1
			end if
		else
		    done = 2
		end if
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
 else
    cmdView	= request("cmdView")
 End if

 Dim rsPurchase
 set rsPurchase = server.CreateObject("Adodb.recordset")
 q="select * from Issue  " & _
  " where SivCode='" & fcCode & "' and " & _
  " fcBA='" & session("fcBA") & "'"
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
    IssueDate=rsPurchase("SivDate")
	xCode=rsPurchase("StoreCode")
 End if
 if done = 1 then 
	connStr.RollbackTrans   
	response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>Nilai Quantity tidak boleh Negatif atau Saldo belum didefinisikan</font>"
 end if
 if done = 2 then 
	connStr.RollbackTrans   
	response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>Terdapat transaksi setelah data ini, harap menghapus transaksi setelah ini</font>"
 end if
 if err.number <> 0 then
   connStr.RollbackTrans   
   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
 else
     if done = 0 then
         connStr.CommitTrans	 
	 end if
 End if
 '-------------ISI Payment Form
 %>
 <form name="frmEntry" method="post" action="IssueDetails.Asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">Issue No</td>
		<td align="left"><input type="text" maxlength="15" name="fcCode" size="30" value="<%=fcCode%>"><font color="#FF0000">*</font></td></tr><tr>
		<td align="left">Date</td>
		<td align="left"><input type="text" maxlength="10" name="IssueDate" size="10" value="<%=IssueDate%>">
		<a href="javascript:show_calendar('frmEntry.IssueDate');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
		</td></tr>
		<tr>
		<td align="left">Store</td>
		<td align="left">      
		<select name="xCode">
	    <%cmdCallCombo "Store",xCode,"",""%>
        </select><font color="#FF0000">*</font>
        </td>
		</td>
		</tr>
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;&nbsp;
	    <a href="javascript:launchwin('PrintSIV.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>&IssueDate=<%=IssueDate%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
		<a href="IssueDetails.Asp?fcCode=<%=fcCode%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" AND LEN(TRIM(fcCode)) > 4 Then
   Dim rsPurchaseOrder,intRecordsPerPage,cColor
   set rsPurchaseOrder = server.CreateObject("Adodb.recordset")
   q="select * from IssueDetails  " & _
     " where SivCode='" & fcCode & "' and " & _
     " fcBA='" & session("fcBA") & "'"
   'response.Write q	
   Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
   rsPurchaseOrder.CursorType = 3
   rsPurchaseOrder.CursorLocation = 3
   rsPurchaseOrder.Open q, connstr%>
   <form action="IssueDetails.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		<tr><td align="left">Item Code</td>
		<td align="left">
		<%IF txtJenis="Edit" Then%>
		     <input type="text" size="15" name="ITEMCODE" value=<%=ITEMCODE%> readonly="True" style="background:#999999">&nbsp;<font color="#FF0000">* <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Stock&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Stock</font></a>
			<tr><td align="left">Job Code</td><td align="left"><input type="text" size="15" name="Jobcode" value=<%=Jobcode%> readonly="True" style="background:#999999">&nbsp;<font color="#FF0000">*
			<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Job</font></a></td>
			</tr>
			<tr><td align="left">TT</td><td align="left"><input type="text" size="15" name="LocationType" value=<%=LocationType%> readonly="True" style="background:#999999">&nbsp;<font color="#FF0000">*<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1"> Location Type</a></td></tr>
			<tr><td align="left">TC</td><td align="left"><input type="text" size="15" name="LocationCode" value=<%=LocationCode%> readonly="True" style="background:#999999">&nbsp;<font color="#FF0000">* </td></tr>


			
			<input type="hidden" size="15" name="LineNoKey" value="<%=LineNoKey%>" readonly="True" style="background:#999999">
			<input type="hidden" size="15" name="Linenohist" value="<%=Linenohist%>" readonly="True" style="background:#999999">
			
		<%else%>
		    <input type="text" size="15" name="ITEMCODE" value=<%=ITEMCODE%>>&nbsp;<font color="#FF0000">* <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Stock&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Stock</font></a>
			<tr><td align="left">Job Code</td><td align="left"><input type="text" size="15" name="Jobcode" value=<%=Jobcode%>>&nbsp;<font color="#FF0000">*
			<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Job</font></a></td>
			</tr>
			<tr><td align="left">TT</td><td align="left"><input type="text" size="15" name="LocationType" value=<%=LocationType%>>&nbsp;<font color="#FF0000">*<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1"> Location Type</a></td></tr>
			
			<tr><td align="left">TC</td><td align="left"><input type="text" size="15" name="LocationCode" value=<%=LocationCode%>>&nbsp;<font color="#FF0000">*</td></tr>
			
			<input type="hidden" size="15" name="LineNoKey" value="<%=LineNoKey%>" readonly="True" style="background:#999999">
			<input type="hidden" size="15" name="Linenohist" value="<%=Linenohist%>" readonly="True" style="background:#999999">
		<%End if%>
		</tr>
		<tr><td align="left">Quantity</td><td align="left"><input type="text" size="15" name="QUANTITY" value=<%=QUANTITY%>>&nbsp;<font color="#FF0000">*</font></td></tr>
		<tr><td align="left">REFERENCE</td><td align="left"><input type="text" size="30" maxlength = "30" name="REFERENCE" value="<%=REFERENCE%>"></td></tr>
		<tr><td align="left">REMARKS</td><td align="left"><textarea  name="REMARKS" maxlength = "100" cols="50" rows="3"><%=REMARKS%></textarea></td></tr>
	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
		  <input type="hidden" name="IssueDate" value="<%=IssueDate%>">
		  <input type="hidden" name="xCode" value="<%=xCode%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: 
		       bold; font-size: 8pt; font-family:Verdana" name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: 
		       bold; font-size: 8pt; font-family:Verdana" name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset" style="background-color: #298CF7; color: #F7C610; font-weight:
		  bold; font-size: 8pt; font-family:Verdana" name="Reset" value="Reset">
		  <input Type="Button" style="background-color: #298CF7; color: #F7C610; font-weight: 
		  bold; font-size: 8pt; font-family:Verdana" value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <table class="data">
	<BR>
	<tr bgcolor="#6666FF" >
	<TD align="left">NO</TD>
	<TD align="left">LineNoKey</TD>
	<TD colspan="2">&nbsp;</TD>	
	<TD align="left">ISSUE CODE</TD>	
	<TD align="left">STOCK CODE</TD>	
	<TD align="left">LOC. CODE</TD>	
	<TD align="left">LOC. TYPE</TD>	
	<TD align="left">QUANTITY</TD>	
	<TD align="left">PRICE</TD>	
	<TD align="left">AMOUNT</TD>	
	<TD align="left">JOB CODE</TD>	
	<TD align="left">Line No (history)</TD>		
	<TD align="left">Reference</TD>	
	<TD align="left">Remarks</TD>	
	</tr>   
   <%
   if Not rsPurchaseOrder.eof then
      nCountofField=rsPurchaseOrder.fields.count-1
	  no=1
	  While Not rsPurchaseOrder.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
		  <td align="left" valign="top"><%=rsPurchaseOrder("LineNoKey")%></td>
		  
			<td valign="middle" width="30">
			<a href="IssueDetails.Asp?fcCode=<%=rsPurchaseOrder("SivCode")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&ItemCode=<%=rsPurchaseOrder("StockCode")%>&Quantity=<%=rsPurchaseOrder("Quantity")%>&JobCode=<%=rsPurchaseOrder("JobCode")%>&LocationType=<%=rsPurchaseOrder("LocationType")%>&LocationCode=<%=rsPurchaseOrder("LocationCode")%>&LineNoHist=<%=rsPurchaseOrder("LineNoHist")%>&Reference=<%=rsPurchaseOrder("Reference")%>&Remarks=<%=rsPurchaseOrder("Remarks")%>&txtJenis=Edit">
			<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsPurchaseOrder(1)%>" border="0"></a></td>
			<td valign="middle" width="10"><a href="IssueDetails.Asp?fcCode=<%=rsPurchaseOrder("SivCode")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&ItemCode=<%=rsPurchaseOrder("StockCode")%>&LineNoHist=<%=rsPurchaseOrder("LineNoHist")%>&txtJenis=Details">
			<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsPurchaseOrder(1)%>" border="0" onClick="return confirm('Are you sure to delete');"></a></td>
		    <%for i = 0 to nCountofField-8%>
			    <td align="left" title="<%=rsPurchaseOrder(i).Name%>" valign="top"><%=rsPurchaseOrder(i)%></td>
		    <%next%>
			<td align="left" title="<%=rsPurchaseOrder("linenohist")%>" valign="top"><%=rsPurchaseOrder("linenohist")%></td>		    
			<td align="left" title="<%=rsPurchaseOrder("Reference")%>" valign="top"><%=rsPurchaseOrder("Reference")%></td>		    
			<td align="left" title="<%=rsPurchaseOrder("Remarks")%>" valign="top"><%=rsPurchaseOrder("Remarks")%></td>		    
		  </tr>
		  <%
		  rsPurchaseOrder.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsPurchaseOrder.Close
   set rsPurchaseOrder = Nothing%>
   </TABLE>
   </form>
<%End if%>  
<%'-------------akhir----------------------%>
</td>
</tr>
<tr>
  <td colspan="2">
  <!--#include file="bottom.jsp"-->
  </td>
</tr>
</table>
</body>
</html>
