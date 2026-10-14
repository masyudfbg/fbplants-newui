<!--#include file="../connDb1.Inc"-->
<%

ajaxUoM = Request("ajaxUoM")

if ajaxUoM = "quality" then

	Dim txtUoM 
	
	response.expires=-1
	 stringAta = " select Stock_Api.Get_UoM('" & request.querystring("q") & "','" & Session("fcBA") & "') tempUoM from dual "

					  Set rsUoM = Server.CreateObject("ADODB.Recordset")

					  rsUoM.Open stringAta, connstr

						   

					  IF Not rsUoM.eof then
						 txtUoM = rsUoM("tempUoM")
					  End if

					  rsUoM.Close

					  set rsUoM = Nothing  
					  
					  
	response.expires=-1
					  
	'get the q parameter from URL
	q=ucase(request.querystring("q"))


	'Output "no suggestion" if no hint were found
	'or output the correct values

	  response.write("&nbsp;"+txtUoM)


End if



if ajaxUoM = "StockName" then

	Dim txtStockName
	
	response.expires=-1
	 stringStockName = " select Stock_Api.Get_stockName('" & request.querystring("q") & "','" & Session("fcBA") & "') tempStockName from dual "
		
		Set rsUoM = Server.CreateObject("ADODB.Recordset")

		rsUoM.Open stringStockName, connstr

			

		IF Not rsUoM.eof then
			txtStockName = rsUoM("tempStockName")
		End if

		rsUoM.Close

		set rsUoM = Nothing  
					  
					  
	response.expires=-1
					  
	'get the q parameter from URL
	q=ucase(request.querystring("q"))


	'Output "no suggestion" if no hint were found
	'or output the correct values

	  response.write("("+txtStockName+")")


End if

if ajaxUoM = "StockItem" then
	
	Dim txtJmlStock
	
	response.expires=-1
	stringStockName = " select openingbalanceqty from stockledger where fcba = '"&Session("fcBA")&"' and year = '"&request.querystring("thn")&"' and month = '"&request.querystring("bln")&"' and itemcode = '"&request.querystring("q")&"' "
	
	Set rsUoM = Server.CreateObject("ADODB.Recordset")
	
	rsUoM.Open stringStockName, connstr
	response.write rsUoM("openingbalanceqty")


End if


%>
