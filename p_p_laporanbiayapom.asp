<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Monitoring PR</title>
<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
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

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF;font-size:1px"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtINVOICEDATE_From= trim(request("CREATIONDATE_From"))
v_txtINVOICEDATE_To= trim(request("CREATIONDATE_To"))

	dim company,bulan,tahun,filtercompany1,fcba,sqlpt,userid,statusgrn,filterStatusGRN,whereClause,doc_classification,filter_doc

company= trim(request("company"))
fcba= trim(request("businessunit"))
doc_classification= trim(request("document_classification"))

userid = session("CodeName")

 q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
	set rsview = connStr.Execute(q)
	if not rsview.eof then
	   fcbaname = rsview("fcName")
	   CompanyName=rsview("FCCOMPANYNAME")
	   MgrName = rsview("FCMANAGER")
	   KTUName = rsview("FCKTU")
	   CropName=rsview("FCCROP")
	   FcbaType=rsview("fcType")
	end if

Dim sqlview,rsview


if ptype <> "-" then 
		filter1 = " AND stock_api.get_stockgroup(st.fccode,st.fcba) ='" & ptype & "' "
end if

if pstore <> "" then 
		filter2 = " AND s.storecode ='" & pstore & "' "
end if


txtSQL = " SELECT " & _ 
       " st.fcba, " & _
       " s.storecode, " & _
       " st.fccode, " & _
       " stock_api.get_stockgroup(st.fccode,st.fcba) StockGroupcode, " & _
       " stockgroup_api.get_stockgrpname(stock_api.get_stockgroup(st.fccode,st.fcba),st.fcba) stockgroupname, " & _
       " nvl(stocksubgroup_api.get_stocksubgrpname(stock_api.get_StockSubGroup(st.fccode,st.fcba)),'-') stocksubgroupname, " & _
       " stock_api.get_ActiveMatrial(st.fccode,st.fcba) ActMat , " & _
       " stock_api.get_stockname(st.fccode, st.fcba) fcname,  " & _
       " stock_api.get_uom(st.fccode, st.fcba) unitofmeasurement,  " & _
       " (Nvl(s.openingbalanceqty,0)) openingbalanceqty, " & _
       " (Nvl(s.qtyreceived,0)) qtyreceived, " & _
       " (Nvl(s.qtyreceivedtask,0)) qtyreceivedtask, " & _
       " (Nvl(s.qtyreceivetaskdiv,0)) qtyreceivetaskdiv, " & _
       " (Nvl(s.qtytransfered,0)) qtytransfered, " & _
       " (Nvl(s.qtytransporttask,0)) qtytransporttask, " & _
       " (Nvl(s.qtytransporttaskdiv,0)) qtytransporttaskdiv, " & _
       " (Nvl(s.qtyreturned,0)) qtyreturned, " & _
       " (Nvl(s.qtyissued,0)) qtyissued, " & _
       " (Nvl(s.qtywriteoff,0)) qtywriteoff, " & _
       " (Nvl(s.endingbalanceqty,0)) endingbalance, " & _
       " Round(Nvl(s.amtperunit,0),2) amtperunit, " & _
       " Decode(s.storecode,'ST1',stock_api.get_stockminimum(st.fccode, st.fcba),0) get_stockminimum, " & _
       " SUM(Nvl(pd.requestedqty,0)) requestedqty, " & _
       " SUM(Nvl(pd.approvedqty,0)) approvedqty, " & _
       " SUM(Nvl(pd.orderedqty,0)) orderedty " & _
       " FROM " & _
       " stock st " & _
       " Left Join stockledger s on st.FCCODE=s.ITEMCODE and st.fcba=s.fcba" & _
       " Left Join PURCHASEREQUESTDETAILS pd on s.ITEMCODE = pd.ITEMCODE and s.FCBA = pd.FCBA " & _
       " WHERE 1=1  " & _
       " AND s.YEAR='" & pyear & "' AND s.MONTH='" & pmonth & "' AND st.FCBA ='" & pfcba & "' " & _
	   filter1 & _
	   filter2 & _
	   " GROUP BY  st.fcba, s.storecode, st.fccode, stock_api.get_stockgroup(st.fccode,st.fcba) , " & _
	   " stockgroup_api.get_stockgrpname(stock_api.get_stockgroup(st.fccode,st.fcba),st.fcba) , " & _
       " nvl(stocksubgroup_api.get_stocksubgrpname(stock_api.get_StockSubGroup(st.fccode,st.fcba)),'-') , " & _
       " stock_api.get_ActiveMatrial(st.fccode,st.fcba)  , stock_api.get_stockname(st.fccode, st.fcba) , " & _
       " stock_api.get_uom(st.fccode, st.fcba),Round(Nvl(s.amtperunit,0),2), " & _
	   " (Nvl(s.openingbalanceqty,0)), " & _
       " (Nvl(s.qtyreceived,0)), " & _
       " (Nvl(s.qtyreceivedtask,0)), " & _
       " (Nvl(s.qtyreceivetaskdiv,0)), " & _
       " (Nvl(s.qtytransfered,0)), " & _
       " (Nvl(s.qtytransporttask,0)), " & _
       " (Nvl(s.qtytransporttaskdiv,0)), " & _
       " (Nvl(s.qtyreturned,0)), " & _
       " (Nvl(s.qtyissued,0)), " & _
       " (Nvl(s.qtywriteoff,0)), " & _
       " (Nvl(s.endingbalanceqty,0)) " & _
       " Order by st.fcba,st.fccode,s.storecode"

txtSelect = pmonth & "/" & pyear

if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'response.Write txtsql
'response.Flush()
%>
	<body>
		<div id="wrapper">
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="28" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
				</div>
				<div id="subheader" class="centered" class="centered" style="width:50%; margin:auto;"> 
					<h2>LAPORAN BIAYA POM</h2>
					<br/>
						<h4>
							PERIODE S/D &nbsp;:&nbsp;<%=textDate(v_txtINVOICEDATE_From)%>  TO <%=textDate(v_txtINVOICEDATE_To)%>
						</h4>
				</div>
			</div>
				</th>
			</tr>
			

				<!--<div class="printhide">-->
						<tr>
							<th colspan="28" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
						<tr>
						
						
						</tr> 
						<tr >
							<tr>
		    <th>No.</th>
			<th>Tanggal</th>
			
			<th>Akun</th>
			<th>Description</th>
			<th>Debet</th>
			<th>Kredit</th>
			<th>Saldo</th>
		
		</tr>
                
                        </tr>
                
                       
                         </thead>

						
						</table>
		</DIV>
	</BODY>
</HTML>	

 

<script type="text/javascript">
function getKeterangan(no) {
  
        var keterangan = $('#keterangan' + no).val();
        var prcode = $('#prcode' + no).val();
        var lineid = $('#lineid' + no).val();
        var itemcode = $('#itemcode' + no).val();
	 	var fcba = $('#fcba' + no).val();
	 	var pocode = $('#pocode' + no).val();
	 	var linenokeypod = $('#linenokeypod' + no).val();

        $.ajax({
            url: "monitoringprproses.asp",
            type: "GET",
            data: {
                keterangan: keterangan,
                prcode: prcode,
                lineid: lineid,
                itemcode: itemcode,
                fcba: fcba,  
                pocode: pocode,  
                linenokeypod: linenokeypod  
            },
			success: function(response) {
            // Menampilkan pesan jika berhasil
            alert("keterangan Berhasil diperbaharui!");
			//location.reload();
        },
        error: function(xhr, status, error) {
            // Menampilkan pesan jika terjadi error
            alert("Gagal memperbaharui data: " + error);
        }
        });
   
}


	function numformat(val) {
		var nStr = Math.round(val).toString();
		nStr = nStr.replace(/\,/g, "");
		var x = nStr.split('.');
		var x1 = x[0];
		var x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) {
			x1 = x1.replace(rgx, '$1' + '.' + '$2');
		}
		return x1 + x2;
	}
	
	function getExport()
	{
		var pfcbagl    = $("#businessunitnc").val();
		var nm = 'Laporan_BiayaPOM' + '_'+tahun+".xlsx";  
	
		TableToExcel.convert(document.getElementById("product-table"), {
            name: nm,
            sheet: {
            name: "Laporan_BiayaPOM"
            }
          });
		//$('table').tblToExcel();
	}



	</script>