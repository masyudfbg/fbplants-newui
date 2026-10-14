<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - UPLOADED TIMBANGAN</title>
<link rel="StyleSheet" href="../style-content.css" type="text/css" />
<link rel="stylesheet" href="../css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="../css/jquery-ui.css" />


<style>
#view p{margin-left:10px;}
#view{display:table-cell;vertical-align:bottom; position:absolute; bottom:0px;background:#000000; color:#fff; margin-bottom:10px;}
</style>
<style>
*{font-size:10px;}
.responded {
    color:green;
}
.checking,.unchecked {
    color:#FF8C00;
}
.timeout {
    color:red;
}
.conn_ip_form{
padding:10px;

}
ul {margin-top:20px;}
ul li {list-style-type: none;margin-bottom:10px;}

.warp_view_01{position:relative; 
background-color:#000;
border:1px #000 solid;
padding-top:10px;
padding-left:0px;
padding-bottom:20px;
width:99.9%; 
min-height:120px;}

.warp_view{
position:relative; 
background-color:#FAFAFA;
border:1px #ccc solid;
padding-top:10px;
padding-bottom:20px;
width:99.8%; 
min-height:100px;
}

</style>
<style>
.blink_me {
-webkit-animation-name: blinker;
-webkit-animation-duration: 1s;
-webkit-animation-timing-function: linear;
-webkit-animation-iteration-count: infinite;

-moz-animation-name: blinker;
-moz-animation-duration: 1s;
-moz-animation-timing-function: linear;
-moz-animation-iteration-count: infinite;

animation-name: blinker;
animation-duration: 1s;
animation-timing-function: linear;
animation-iteration-count: infinite;
}
.alert {
-webkit-animation-name: blinker;
-webkit-animation-duration: 7s;
-webkit-animation-timing-function: linear;
-webkit-animation-iteration-count: infinite;

-moz-animation-name: blinker;
-moz-animation-duration: 1s;
-moz-animation-timing-function: linear;
-moz-animation-iteration-count: infinite;

animation-name: blinker;
animation-duration: 1s;
animation-timing-function: linear;
animation-iteration-count: infinite;
}
@-moz-keyframes blinker {  
0% { opacity: 1.0; }
50% { opacity: 0.0; }
100% { opacity: 1.0; }
}

@-webkit-keyframes blinker {  
0% { opacity: 1.0; }
50% { opacity: 0.0; }
100% { opacity: 1.0; }
}
@keyframes blinker {  
0% { opacity: 1.0; }
50% { opacity: 0.0; }
100% { opacity: 1.0; }
}
</style>
<!-- For ease i'm just using a JQuery version hosted by JQuery- you can download any version and link to it locally -->
<script src="../ajax/knockout-min.js"></script>
<script src="../js/jquery-1.9.1.js"></script>
<script>
(function($)
{
    $(document).ready(function()
    {
        
		
        var $container = $("#data");
        $container.load("../ajax/ExportDataWbs_Ajax.asp");
        var refreshId = setInterval(
		function(){
            $container.load('../ajax/ExportDataWbs_Ajax.asp');
        }, 15000);
		
		var $container_ = $("#view");
        $container_.load("../ajax/ViewDataExportToIplas.asp");
        var refreshId_ = setInterval(
		function(){
            $container_.load('../ajax/ViewDataExportToIplas.asp');
        }, 15000);
    });
})(jQuery);

</script>
<script>
//<![CDATA[

function ping(ip, callback) {

    if (!this.inUse) {
        this.status = 'unchecked';
        this.inUse = true;
        this.callback = callback;
        this.ip = ip;
        var _that = this;
        this.img = new Image();
        this.img.onload = function () {
            _that.inUse = false;
            _that.callback('responded');

        };
        this.img.onerror = function (e) {
            if (_that.inUse) {
                _that.inUse = false;
                _that.callback('responded', e);
            }

        };
        this.start = new Date().getTime();
        this.img.src = "http://" + ip;
        this.timer = setTimeout(function () {
            if (_that.inUse) {
                _that.inUse = false;
                _that.callback('timeout');
            }
        }, 1500);
    }
}

var PingModel = function (servers) {

    var self = this;
    var myServers = [];
	var iplas = "172.16.100.1";
    ko.utils.arrayForEach(servers, function (location) {

        myServers.push({
            name: location,
            status: ko.observable('unchecked')
			
        });
		
    });
	
    self.servers = ko.observableArray(myServers);
    ko.utils.arrayForEach(self.servers(), function (s) {
		if (s.name == "IPLAS"){
		names = "172.16.100.1"
		}else if (s.name == "NULL"){
		names="NULL"
		}else{
		names = s.name
		}
        s.status('checking');
        new ping(names, function (status, e) {
            s.status(status);
        });
		
    });
};
window.onload=function(){
var ip001 = document.getElementById('ip001').value;
var ip002 = document.getElementById('ip002').value;
var ip003 = document.getElementById('ip003').value;
if (ip001==""){
ip001="NULL"
}
if (ip002==""){
ip002="NULL"
}
if (ip003==""){
ip003="NULL"
}
var komodel = new PingModel(['IPLAS',
	ip001,
	ip002,
	ip003,
    ]);
	
ko.applyBindings(komodel)
}//]]> 
</script>

</head>

	<body>
			<div id="wrapper">
						<table border="0" width="100%" id="product-table" class="xthf-site">
							 <thead>
								 <tr>
									<th colspan="10" style=" background-color:#FFFFFF; color:#000000; border:none">
									
											<h2>MONITORING UPLOADED DATA TIMBANGAN TO IPLAS</h2>

									</th>
								</tr>
								<tr>
									 
									<th WIDTH="20">NO</th>
									<th>FCBA</th>
									<th>ERROR MESSAGE</th>
									<th>DATE</th>

								</tr>
							</thead>
							<tbody id="data" >
								<tr STYLE="BACKGROUND-COLOR:#FAFAFA;">
									<td colspan="4" align="CENTER">PROCESSING UPLOAD DATA <BR/><img src="../IMAGES/loader2.gif"> </td>
								</tr>
							</tbody>   
						</table>

				<div class="warp_view_01">
					<div id="view" ></div>

				</div>		
				<div style="background-color:#FAFAFA;border:1px #ccc solid;padding-top:10px;padding-bottom:20px;">
					<div style="text-align:center;">PROCESSING UPLOAD DATA<BR/><img src="../IMAGES/loader2.gif"></div>
				</div>
				<div class="warp_view">
					<div class="conn_ip_form">
						<h3>Please, Check Connection (IP)</h3>
						<br/>
						<input id="ip001" type="text" value="" width="200">
						<input id="ip002" type="text" value="" width="200">
						<input id="ip003" type="text" value="" width="200">
						<a href="#" onclick="location.reload(); " class="button_menu_top" style="color:#fff;">Check Connection </a>
							<ul data-bind="foreach:servers">
								<li>
									<a href="#" data-bind="text:name,attr:{href: 'http://'+name}">tester</a>
									<span data-bind="text:status,css:status"></span>
								</li>
							</ul>
					</div>
				</div>		
		</div>
		

    <div id="content"></div>
    <div id="loading"></div>
 
</body>
</html>		
