<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<html>
<head>
<title>ePlaS - Upload Manual Fingerprint</title>
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

</head>
<body>

<!--#include file="top.jsp"-->

<div style="width:21.7%; float:left; margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
 <!--#include file="menu.asp"-->
</div>

<div class="judul_head" valign="top" bgcolor="#fff" 
     style="width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 

    <div class="judul">
       Module : <%=request("txtModule")%> / Upload Stock Opname
    </div>

 <div class="box_both" style="overflow:scroll;"> 

 <h3>Upload Stock Opname</h3>

 <!--================ FORM UPLOAD =================-->
 <form name="frmEntry" method="post" action="simpan_upload_stockopname.asp" enctype="multipart/form-data">
	<table class="data">
		<tr>
			<td align="left"><span style="font-size:14px">Upload File (.xlsx)</span></td>
			<td align="left">
                <input type="file" name="fileupload" required>
                <font color="#FF0000">*</font>
            </td>
		</tr>

        <tr>
            <td></td>
            <td>
                <input type="submit" value="Upload" class="button">
                <input type="reset" value="Reset" class="button">
            </td>
        </tr>
    </table>
 </form>
 <!--============== END FORM =================-->

 <!--#include file="bottom.jsp"-->

</div>
</div>

</body>
</html>
