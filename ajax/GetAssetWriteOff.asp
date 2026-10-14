<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
itemcode = Request("itemcode")
result = ""

q =   " SELECT DISTINCT ASH.PO_CODE, ASH.ACQUISITIONVALUE, ASH.CURRENCYCODE, ASH.ITEMDESCRIPTION, ASH.QTY_HDR, ASH.RECEIVENOTECODE, ASH.RECEIVENOTEDATE," & _
	  " ASH.SUPPLIERNAME, ASD.ADDITIONVALUE, ASD.AMOUNT, ASD.AMOUNTDEPRECIATION, ASD.ASSETNAME, ASD.BRAND, ASD.COSTCENTER, ASD.ESTIMATELIFEYEAR, ASD.FA_GROUPID, " & _
	  "	ASD.FA_GROUPNAME, ASD.JOBCODE, ASD.LOCATION_ID, ASD.MADEIN, ASD.QTY_DTL, ASD.RESIDUALVALUE, ASD.SERIALNUMBER, ASD.SPECIFICATION, ASD.STARTDEPRECIATION, " & _
	  "	ASD.TAGNUMBER, ASD.VOUCHERNUMBER, ASD.INSTALLDATE, ASH.CURRENCYRATE " & _
	  "	FROM ASSETADDITIONALHDR ASH " & _
	  "	INNER JOIN ASSETADDITIONALDTL ASD ON (ASH.ITEMCODE = ASD.ITEMCODE AND ASH.FCBA =  ASD.FCBA) " & _
	  "	WHERE ASH.FCBA = '"& Session("FCBA") &"'" & _
	  "	AND ASH.ITEMCODE = '"& itemcode &"'"


' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	assetname = rsrecurtrans("assetname")
	fa_groupid = rsrecurtrans("fa_groupid")
	acquisitionvalue = rsrecurtrans("acquisitionvalue")
	costcenter = rsrecurtrans("costcenter")
	location_id = rsrecurtrans("location_id")
	qty_dtl = rsrecurtrans("qty_dtl")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """assetname"" : """ & assetname & ""","
response.write """fa_groupid"" : """ & fa_groupid & ""","
response.write """acquisitionvalue"" : """ & acquisitionvalue & ""","
response.write """costcenter"" : """ & costcenter & ""","
response.write """location_id"" : """ & location_id & ""","
response.write """qty_dtl"" : """ & qty_dtl & """"
response.write "}"
%>