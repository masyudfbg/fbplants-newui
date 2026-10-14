<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Purchase Order Edit</title>
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

<link href="plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="plugins/responsivegridsystem/12cols.css" media="all">


</head>
<body>
<!--#include file="top.jsp"-->
<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
 <!--#include file="menu.asp"-->
</div>
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
    <div class="judul">
       Purchase Order Edit Header
    </div>
 <div class="box_both" style="overflow:scroll;"> 
 <%
  Dim txtLineNoKey : txtLineNoKey = Request("txtLineNoKey")

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")


 fcCode = Request("fcCode")
 'HeaderRemarks = Request("HeaderRemarks")
 businessunit = Request("businessunit")

 dim sqldata,rsheader
 dim tglpo,purchaseby,incoterms,clasification,remarks,termofpayment,dp,ppnptg,ppnamt,suppliercode
 dim ttlamt,ttlamtppn,ttlcharge,grandtotal,documentname
 dim tglporeq,purchasebyreq,incotermsreq,clasificationreq,remarksreq,termofpaymentreq,dpreq,ppnptgreq,ppnamtreq
 dim ttlamtreq,ttlamtppnreq,ttlchargereq,grandtotalreq,documentnamereq
 sqldata = "Select H.POCODE, TO_CHAR(H.PODATE,'MM/DD/YYYY')PODATE, H.SUPPLIERCODE, H.REMARKS, H.FCENTRY, " & _
      " H.FCEDIT, H.FCIP, H.FCBA, H.ROWSTATEID, H.LASTUPDATE, " & _
      " H.LASTTIME, H.USE_PPN, H.ISINBRIDGE, H.INCOTERM1, H.INCOTERM2,  " & _
      " H.PPN, H.TOTALAMOUNTPPN, H.COMPANY, H.SUPPLIERADDRESS, H.PURCHASEBY,  " & _
      " H.DELIVERYDESTINATION, H.TERMOFPAYMENT, H.DP, H.TOTALCHARGE, H.GRANDTOTAL,  " & _
      " H.TOTALAMOUNT, H.TOTALPPN, H.DOCUMENT_CLASSIFICATION,dc.fcname, H.CURRENCYCODE, H.EXCHANGERATE,  " & _
      " H.POTYPE " & _
" From  PURCHASEORDER H INNER JOIN document_classification dc ON H.DOCUMENT_CLASSIFICATION=dc.fccode " & _
" INNER JOIN supplier S ON S.FCCODE= H.SUPPLIERCODE" & _
" WHERE H.POCODE='" & Request("fcCode") & "' AND dc.document_type= 'PO' and dc.ISACTIVE='TRUE'"
set rsheader = connStr.Execute(sqldata)
'response.write sqldata
IF  not rsheader.eof then 
    tglpo           = rsheader("PODATE")
    purchaseby      = rsheader("PURCHASEBY")
    incoterms       = rsheader("INCOTERM1")
    clasification   = rsheader("DOCUMENT_CLASSIFICATION")
    documentname    = rsheader("fcname")
    remarks         = rsheader("REMARKS")
    termofpayment   = rsheader("TERMOFPAYMENT")
    dp              = rsheader("DP")
    ppnptg           = rsheader("PPN")
    ppnamt           = rsheader("TOTALPPN")
    ttlamt           = rsheader("TOTALAMOUNT")
    ttlamtppn        = rsheader("TOTALAMOUNTPPN")
    ttlcharge        = rsheader("TOTALCHARGE")
    grandtotal       = rsheader("GRANDTOTAL")
	supplieraddress  = rsheader("SUPPLIERADDRESS")
	grandtotal       = rsheader("GRANDTOTAL")
	suppliercode     = rsheader("SUPPLIERCODE")
   

end if
 'response.write fcCode & "-" & tglpo

tglporeq           = Request("PODATE")
purchasebyreq      = Request("PURCHASEBY")
incotermsreq       = Request("SHIPPINGPOINT")
clasificationreq   = Request("DOCUMENT_CLASSIFICATION")
remarksreq         = Request("HeaderRemarks")
termofpaymentreq   = Request("TERMOFPAYMENT")
dpreq              = Request("DP")
ppnptgreq           = Request("PPN")
ppnamtreq           = Request("TOTALPPN")
ttlamtreq           = Request("TOTALAMOUNT")
ttlamtppnreq        = Request("TOTALAMOUNTPPN")
ttlchargereq        = Request("TOTALCHARGE")
grandtotalreq       = Request("GRANDTOTAL")
suppliercodereq 	= Request("SUPPLIERCODE")
suppliername 		= Request("suppliername")
supplieraddressreq = Request("supplieraddress")

 '--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")= "Update"   Then

    if Request("txtJenis")="Update" then
    ' replace('" & totalcharge & "',',','')
       q =  " Update purchaseorder" & _
            " set podate = to_date('" & tglporeq & "','mm/dd/yyyy') " & _ 
            "     , PURCHASEBY = '" & purchasebyreq & "'" & _
            "     , INCOTERM1 = '" & incotermsreq & "'" & _
            "     , DOCUMENT_CLASSIFICATION = '" & clasificationreq & "'" & _
            "     , Remarks = '" & remarksreq & "'" & _
			"     , supplieraddress 	='" & supplieraddressreq & "'" & _
			"     , suppliercode 	='" & suppliercodereq & "'" & _
            "     , TERMOFPAYMENT = '" & termofpaymentreq & "'" & _
            "     , dp = replace('" & dpreq & "',',','') " & _
            "     , PPN = replace('" & ppnptgreq & "',',','') " & _
            "     , TOTALPPN = replace('" & ppnamtreq & "',',','') " & _
            "     , TOTALAMOUNT = replace('" & ttlamtreq & "',',','') " & _
            "     , TOTALAMOUNTPPN = replace('" & ttlamtppnreq & "',',','') " & _
            "     , TOTALCHARGE = replace('" & ttlchargereq & "',',','') " & _
            "     , GRANDTOTAL = replace('" & grandtotalreq & "',',','') " & _
            "     , fcedit = '" & Session("CodeName") & "'" & _
            "     , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
            "     , lastupdate = sysdate " & _
            "     , lasttime = to_char(sysdate,'hh24:mi') " & _
            "       where pocode = '" & Request("fcCode") & "'" & _
            "       and fcBA = '" & Request("businessunit") & "'"
        ' response.write q
        connStr.execute(q)
         response.redirect("purchaseorderheaderedit.asp?fcCode="&fcCode)
     end if                              
        cmdView = "True"   
else
        cmdView = request("cmdView")
End if


 
  'response.write "IsInBridge--> " & IsInBridge
if err.number <> 0 then
    connStr.RollbackTrans
    response.Write  "<div class='error'>" & err.description & "</div>"
End if

 '-------------ISI Payment Form
%>
 <form name="frmEntry" method="post" action="PurchaseOrderHeaderEdit.Asp">
    <table class="data" width="100%">
        <tr>
            <td align="left">
                PO Code
            </td>
            <td align="left">
                <input type="text"  name="fccode" size="30" value="<%=fcCode%>" style="background:#eee" readonly>
            </td>
         			
				<td align="left">Term Of Payment</td>
				<td align="left">
					<select name="termofpayment" value="<%=termofpayment%>" class="chosen-select" style="width: 40%">
						<option value="<%=termofpayment%>"><%=termofpayment%></option> 
							<% 	q = " SELECT DISTINCT fccode, fcname FROM parameter WHERE parcode='TERM_OF_PAYMENT' ORDER BY fccode, fcname"
									' response.write q
									Set rsComboP = Server.CreateObject("ADODB.Recordset")
									rsComboP.CursorType = 3
									rsComboP.CursorLocation = 3
									rsComboP.Open q, connstr
									if not rsComboP.eof then%>
									<%
									Do While Not rsComboP.eof
									%>
										<option value="<%=rsComboP("fccode")%>" 
											<%if cstr(""&termofpayment)=cstr(""&rsComboP("fccode")) then %> selected 
											<%End If%>><%=rsComboP("fcname")%> 
										</option> 
									<%rsComboP.movenext
									loop
									End if
									rsComboP.Close
									set rsComboP = Nothing
									q=""
							%>
					</select>
 				</td>
        </tr>
		<tr>
            <td align="left">
                Business Unit
            </td>
            <td align="left">
                <input type="text"  name="businessunit" size="30" value="<%=businessunit%>" style="background:#eee" readonly>
            </td>
            <td align="left">DP</td>
				<td align="left">
					<input type="text" maxlength="50" name="dp" size="30" value="<%=dp%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">
				</td>
        </tr>
        <tr>
               <td align="left">Date</td>
				<td align="left">
					<input class="datepicker" type="text"  maxlength="10" name="PODATE" size="10" value="<%=tglpo%>">
			    </td>
                <td align="left">PPn (%)</td>
					<td align="left">
						<input type="text" size="5" name="ppn" id="ppn"  value="<%=ppnptg%>" onkeypress="return harusangka(event);" onkeyup="hitungPPN();">
						<input type="text" size="18" name="totalppn" id="totalppn" value="<%=ppnamt%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" >
				</td>
        </tr>
        <tr>
			<td align="left">Purchase By</td>
				<td align="left">      
					<select name="purchaseby" id="purchaseby" style="width: 70%">
					<option value="<%=purchaseby%>"><%=purchaseby%></option> 
						<% 	Dim rsComboC
							q = " SELECT DISTINCT fccompanycode " & _
							 	" FROM (SELECT businessunit.fccompanycode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"')"
							' response.write q
							Set rsComboC = Server.CreateObject("ADODB.Recordset")
							rsComboC.CursorType = 3
							rsComboC.CursorLocation = 3
							rsComboC.Open q, connstr
							if not rsComboC.eof then%>
							<%
							Do While Not rsComboC.eof
							%>
								<option value="<%=rsComboC("fccompanycode")%>" 
									<%if cstr(""&purchaseby)=cstr(""&rsComboC("fccompanycode")) then %> selected 
									<%End If%>><%=rsComboC("fccompanycode")%>
								</option> 
							<%rsComboC.movenext
							loop
							End if
							rsComboC.Close
							set rsComboC = Nothing
							q=""
						%>
				</select> 
				</td>
			<td align="left">Total Amount</td>
					<td align="left">
						<input type="text" maxlength="50" name="totalamount" id="totalamount" size="15" value="<%=ttlamt%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
						<input type="hidden" maxlength="50" name="totalamounthdn"  id="totalamounthdn" size="15" value="<%=ttlamt%>" style="background:#eee" readonly>
					</td>
		</tr>
		<tr>
			<td align="left">Inco Terms</td>	
				<td align="left">
					<select name="SHIPPINGPOINT" value="<%=incoterms%>" class="chosen-select" style="width: 70%">
						<option value="<%=incoterms%>"></option>
						<option value="FRC" <% If cstr(""&incoterms) = "FRC" Then %> Selected <%End if%>>FRC</option> 
						<option value="FRT" <% If cstr(""&incoterms) = "FRT" Then %> Selected <%End if%>>FRT</option> 
						<option value="LCO" <% If cstr(""&incoterms) = "LCO" Then %> Selected <%End if%>>LCO</option> 
					</select>
				</td>	
				
				<td align="left">Total Amount + PPn</td>
					<td align="left">
						<input type="text" maxlength="50" name="totalamountppn" id="totalamountppn" size="15" value="<%=ttlamtppn%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly >
						
					</td>
			<tr>
				<td align="left">Classification</td>
		<td align="left">
			<select name="document_classification" class="chosen-select" style="width: 70%">
			<option value=<%=clasification%>><%=documentname%></option>
			<%
			qLov = " SELECT fccode, fcname FROM document_classification where document_type= 'PO' and ISACTIVE='TRUE' ORDER BY fccode"
   
			Set rsLov = Server.CreateObject("ADODB.Recordset")
			rsLov.Open qLov, connstr
			if not rsLov.eof then
				while not rsLov.eof
			%>
				<option value="<%=rsLov("fccode")%>" <% if cstr(""&rsLov("fccode")) = cstr(""&document_classification) then%> selected <%end if%>><%=rsLov("fcname")%></option>
			<%
					rsLov.movenext
				wend
			end if
			%>
			</select>
			<font color="#FF0000">*</font>
				<td align="left">Total Charge</td>
				<td align="left"><input type="text" size="15" name="totalcharge" value="<%=ttlcharge%>"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
				</td>	
			</tr>
		</tr>
		<tr>
			<td align="left">Supplier</td>
			<td align="left">
					<select name="suppliercode" value="<%=suppliercode%>" id="suppliercode"  class="chosen-select" style="width: 70%">
						<% 	Dim rsCombosup
							q = " SELECT DISTINCT fccode,fccode || ' - ' || fcname supname " & _
							    " FROM supplier where fcba='"&  Session("fcba") &"' "
							 response.write q
							Set rsCombosup = Server.CreateObject("ADODB.Recordset")
							rsCombosup.CursorType = 3
							rsCombosup.CursorLocation = 3
							rsCombosup.Open q, connstr
							if not rsCombosup.eof then%>
							<%
							Do While Not rsCombosup.eof
							%>
								<option value="<%=rsCombosup("fccode")%>" 
									<%if cstr(""&suppliercode)=cstr(""&rsCombosup("fccode")) then %> selected 
									<%End If%>><%=rsCombosup("supname")%>
								</option> 
							<%rsCombosup.movenext
							loop
							End if
							rsCombosup.Close
							set rsCombosup = Nothing
							
						%>
				</select> 
			<td align="left">Supplier Address</td>
				<td align="left">
						<textarea id="supplieraddress"  name="supplieraddress" style="width: 70%" style="background:#eee" readonly><%=supplieraddress%></textarea>
				</td>
		</tr>
		<tr>
		 <td align="left">
                   Remarks
                </td>
                <td align="left">
                    <textarea name="HeaderRemarks" cols="50" rows="3"><%=remarks%></textarea>
                </td>
		</td>	
			<td align="left">Grand Total</td>
			<td align="left">
				<input type="text" maxlength="50" name="grandtotal" id="grandtotal" size="15" value="<%=grandtotal%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
			</td>	
		</tr>
        
        <tr>
            <td colspan="2"  align="left">
                <input Type="Hidden" Value="True" Name="cmdView">
                <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
                <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
                <input type="submit" name="txtJenis" value="Update" onClick="return confirm('Are you sure to Update this data : <%=fcCode%> ?');"></input>
        </td>
        </tr>
        
    </table>
</div>
  <!--#include file="bottom.jsp"-->
    <script type="text/javascript">
    $(function() {
        $( ".datepicker" ).datepicker();
    });
        function handleClick(cb){
            if(cb.checked == 1){
                cb.value = "Y";
            } else {
                cb.value = "N";
            }
        }

        function hitungPPN()
        {
            var amount                  = parseFloat($('#totalamounthdn').val());
            var ppn                     = parseFloat($('#ppn').val());
            var totalamountppnhdn       = parseFloat($('#totalamountppnhdn').val());
            var ttlppn                  = parseFloat(amount)*(parseFloat(ppn)/100);
            var ttlamountppn            = parseFloat(amount)+parseFloat(ttlppn);
         //   alert(ttlamountppn);
           $('#totalppn').val(ttlppn);
           $('#totalamountppn').val(ttlamountppn);
           $('#grandtotal').val(ttlamountppn);
        }
    </script>
        <script src="/js/jquery.inputmask.bundle.min.js"></script>
        <script src="/js/inputmask.binding.js"></script>

	<script src="/js/jquery.inputmask.bundle.min.js"></script>
<script src="/js/inputmask.binding.js"></script>
<script src="plugins/chosen/chosen.jquery.min.js"></script>
<script type="text/javascript">	
	// $(".chosen-select").empty();
	
	var config = {
	  '.chosen-select'           : {},
	  '.chosen-select-deselect'  : {allow_single_deselect:true},
	  '.chosen-select-no-single' : {disable_search_threshold:10},
	  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
	  '.chosen-select-width'     : {width:"95%"}
	}
	for (var selector in config) {
	  $(selector).chosen(config[selector]);
	}
	// $(".chosen-select").trigger("chosen:updated");

</script>
</body>
</html>
