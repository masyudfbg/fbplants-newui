<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - ISSUE Edit</title>
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
       Issue Edit Header
    </div>
 <div class="box_both" style="overflow:scroll;"> 
 <%
  Dim txtLineNoKey : txtLineNoKey = Request("txtLineNoKey")

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")


 fcCode = Request("fcCode")
 fcba = Request("fcba")

 dim sqldata,rsheader
 dim IssueDate,clasification,xCode,docref,rowstate
 dim IssueDatereq,storecodereq,docrefreq,rowstatereq
 sqldata = "Select I.SIVCODE, I.STORECODE, I.SIVDATE, I.FCENTRY, I.FCEDIT, " & _
      " I.FCIP, I.FCBA, I.LASTUPDATE, I.LASTTIME, I.DOC_REF, " & _
      " I.POSTINGFLG, I.ROWSTATE, I.DOCUMENT_CLASSIFICATION " & _
" From  ISSUE I " & _
" INNER JOIN document_classification dc ON I.DOCUMENT_CLASSIFICATION=dc.fccode " & _
" WHERE I.SIVCODE='" & Request("fcCode") & "' AND I.FCBA='" & Request("fcba") & "' AND dc.document_type= 'SI' and dc.ISACTIVE='TRUE'"

set rsheader = connStr.Execute(sqldata)
'response.write sqldata
IF  not rsheader.eof then 
    IssueDate   = Request("IssueDate")
    clasification   = rsheader("DOCUMENT_CLASSIFICATION")
    xCode       = rsheader("STORECODE")
    docref          = rsheader("DOC_REF")
    rowstate        = rsheader("ROWSTATE")
   

end if
 'response.write fcCode & "-" & tglpo



 '--- Action ke Database di Oracle [ insert,edit,delete ]
'if Request("txtJenis")= "ViewEdit"   Then

if Request("edit")= "Update"   Then
IssueDatereq   = Request("IssueDate")
clasificationreq        = Request("DOCUMENT_CLASSIFICATION")
xCodereq           = Request("STORECODE")
    docrefreq          = Request("DOC_REF")
    rowstatereq        = Request("ROWSTATE")
  
    ' replace('" & totalcharge & "',',','')
       q =  " Update ISSUE" & _
            " set SIVDATE = to_date('" & IssueDatereq & "','mm/dd/yyyy') " & _ 
            "     , fcedit = '" & Session("CodeName") & "'" & _
            "     , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
            "     , lastupdate = sysdate " & _
            "       where SIVCODE = '" & Request("fcCode") & "'" & _
            "       and fcBA = '" & Request("fcba") & "'"
        'response.write (q)
        connStr.execute(q)
       '  response.redirect("issueheader_edit.asp?fcCode="&fcCode)
      '  end if
'else
     '  cmdView = request("cmdView")
End if




 '-------------ISI Payment Form
%>

<form name="frmEntry" method="post" action="issueheader_edit.asp">
	<table class="data" width="100%">
<tr>
			<td align="left">Issue No</td>
			<td align="left"><input type="text" maxlength="30" name="fcCode" size="30" value="<%=fcCode%>" style="background:#eee" readonly>&nbsp;*</font></td>
            <input Type="Hidden" Value="<%=request("fcba")%>" Name="fcba">
            </tr>
		<tr>
			<td align="left">Doc Ref.</td>
			<td align="left"><input type="text" maxlength="50" name="doc_ref" size="30" value="<%=doc_ref%>"></td>
		<tr>
		<td align="left">Date</td>
        <td align="left">
					<input class="datepicker"  onchange="javascript:get_Allowentry();reloadSivCode()" type="text"  maxlength="10" name="IssueDate" size="10"  value="<%=IssueDate%>"><font color="#FF0000">&nbsp;*</font>
			    </td>
        </td></tr>
		<tr>
		<td align="left">Store</td>
		<td align="left">      
		<select name="xCode" id="xcode" onchange="get_Allowentry()">
	    <%cmdCallCombo "Store",xCode,"",""%>
        </select><font color="#FF0000">&nbsp;*</font>
        </td>
		</td>
		</tr>
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Hidden" Value="<%=postingflg%>" id="postingflg_">
		<input Type="Hidden" Value="<%=closeflg_%>" id="closeflg_">

		<input Type="Submit"  name="edit" value="Update">&nbsp;&nbsp;
		
		</td>
		</tr>
    </table>  
</div>
  <!--#include file="bottom.jsp"-->
    <script type="text/javascript">
    $(function() {
        $( ".datepicker" ).datepicker();
        get_Allowentry();
    });
        
        function get_Allowentry()
{
	var tglbkm = $('#IssueDate').val();
	var postingflg_ = $('#postingflg_').val();
	if (tglbkm != "") {
		$.post("bkm_allow_entry.asp?tglbkm=" + tglbkm, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			if(closingStatus=="FALSE"){
				if(postingflg_==1){
					$('#deletehdr').hide();
					$('#poshdr').hide();
				}else{
					$('#add').show();
					$('#eForm').show();
					$('#deletehdr').show();
					$('#poshdr').show();
				}
			}else{
				alert("PERIODE SUDAH DI CLOSE !!!");
				$('#add').hide();
				$('#eForm').hide();
				$('#deletehdr').hide();
				$('#Add').hide();
				$('#poshdr').hide();
			}
			
		});
		
	} 
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
