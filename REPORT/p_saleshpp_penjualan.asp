<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
maxCol = 0
add_col_bud_prod_qty = ""
d_span = 18
	
v_txtFcba = Request("txtFcba")
v_txtYear = Request("txtYear")
v_txtmonth = Request("txtMonth")
v_txtDate = Request("txtDate")

If "" & v_txtDate = "" Then v_txtDate = TextDate(date)

if request("txtYear")="" then
   v_txtYear	=	Year(Session("v_Costbook"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Costbook"))
end if

if  v_txtfcba <> "--ALL--"  and v_txtfcba <> "" then
	whereCons = whereCons & " And fcba = '" & v_txtfcba & "' "
end if
' selectCons = selectCons & ",fcba"
' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
set rsCat = connStr.Execute(q)
if not rsCat.eof then
	fcbaname = rsCat("fcName")
	CompanyName=rsCat("FCCOMPANYNAME")
	MgrName = rsCat("FCMANAGER")
	KTUName = rsCat("FCKTU")
	CropName=rsCat("FCCROP")
	FcbaType=rsCat("fcType")
else
    nSpanTotal = nSpanTotal - 1 
end if

strDefaultBa = 	" SELECT" & _
				" FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
				"     AND FCTYPE = 'M'" & _
				"     AND FCUSERCODE = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("FCBA")
Else
	varDefaultBa = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

arr_field_eof = true
strPusingan	= 	" select " & _
				" fcba," & _
				" product_code," & _
				" month_saldo_awal," & _
				" year_saldo_awal," & _
				" saldo_qty_awal," & _
				" saldo_amount_awal," & _
				" Price_saldo_awal," & _
				" month_prod_sbi," & _
				" year_prod_sbi," & _
				" qty_prod_sbi," & _
				" last_month_qty_prod_sbi," & _
				" amount_prod_sbi," & _
				" last_month_amount_prod_sbi," & _
				" Price_prod_sbi," & _
				" month_qty_sales_sbi," & _
				" year_qty_sales_sbi," & _
				" qty_sales_sbi," & _
				" last_month_qty_sales_sbi," & _
				" (qty_sales_sbi-last_month_qty_sales_sbi) as qty_sales_bi," & _
				" HPP_Qty_level_1," & _
				" HPP_Amount_level_1," & _
				" HPP_Price_level_1," & _
				" HPP_Qty_level_2," & _
				" HPP_Amount_level_2," & _
				" HPP_Price_level_2," & _
				" amount_sales_sbi," & _
				" last_month_amount_sales_sbi," & _
				" (amount_sales_sbi-last_month_amount_sales_sbi) as amount_sales_bi," & _
				" price_sales_sbi," & _
				" case when (qty_sales_sbi-last_month_qty_sales_sbi) <> 0 then ((amount_sales_sbi-last_month_amount_sales_sbi)/(qty_sales_sbi-last_month_qty_sales_sbi)) else 0 end  price_sales_bi," & _
				" saldo_Qty_akhir," & _
				" saldo_amount_akhir," & _
				" price_saldo_akhir" & _
				" from (" & _
				"     select " & _
				"     fcba," & _
				"     product_code," & _
				"     month_saldo_awal," & _
				"     year_saldo_awal," & _
				"     saldo_qty_awal," & _
				"     saldo_amount_awal," & _
				"     Price_saldo_awal," & _
				"     month_prod_sbi," & _
				"     year_prod_sbi," & _
				"     qty_prod_sbi," & _
				"     nvl(IPLASPROD.GetLastMonthqtysbi(product_code, fcba,LAST_DAY(ADD_MONTHS(to_date('"& v_txtMonth &"','fmMM'),-1))),0) last_month_qty_prod_sbi," & _
				"     amount_prod_sbi," & _
				"     nvl(IPLASPROD.GetLastMonthamountprodsbi(product_code, fcba,LAST_DAY(ADD_MONTHS(to_date('"& v_txtMonth &"','fmMM'),-1))),0) last_month_amount_prod_sbi," & _
				"     Price_prod_sbi," & _
				"     month_qty_sales_sbi," & _
				"     year_qty_sales_sbi," & _
				"     qty_sales_sbi," & _
				"     nvl(IPLASPROD.GetLastMonthqtysalessbi(product_code, fcba,LAST_DAY(ADD_MONTHS(to_date('"& v_txtMonth &"','fmMM'),-1))),0) last_month_qty_sales_sbi," & _
				"     HPP_Qty_level_1," & _
				"     HPP_Amount_level_1," & _
				"     HPP_Price_level_1," & _
				"     HPP_Qty_level_2," & _
				"     HPP_Amount_level_2," & _
				"     HPP_Price_level_2," & _
				"     amount_sales_sbi," & _
				"     nvl(IPLASPROD.GetLastMonthamountsalessbi(product_code, fcba,LAST_DAY(ADD_MONTHS(to_date('"& v_txtMonth &"','fmMM'),-1))),0) last_month_amount_sales_sbi," & _
				"     price_sales_sbi," & _
				"     saldo_Qty_akhir," & _
				"     saldo_amount_akhir," & _
				"     price_saldo_akhir" & _
				"     from (" & _
				"         select " & _
				"         fcba," & _
				"         product_code," & _
				"         month_saldo_awal," & _
				"         year_saldo_awal," & _
				"         saldo_qty_awal," & _
				"         saldo_amount_awal," & _
				"         Price_saldo_awal," & _
				"         month_prod_sbi," & _
				"         year_prod_sbi," & _
				"         qty_prod_sbi," & _
				"         amount_prod_sbi," & _
				"         Price_prod_sbi," & _
				"         month_qty_sales_sbi," & _
				"         year_qty_sales_sbi," & _
				"         qty_sales_sbi," & _
				"         HPP_Qty_level_1," & _
				"         HPP_Amount_level_1," & _
				"         HPP_Price_level_1," & _
				"         HPP_Qty_level_2," & _
				"         HPP_Amount_level_2," & _
				"         HPP_Price_level_2," & _
				"         amount_sales_sbi," & _
				"         price_sales_sbi," & _
				"         saldo_Qty_akhir," & _
				"         saldo_amount_akhir," & _
				"         price_saldo_akhir" & _
				"         from (" & _
				"             select" & _
				"             fcba," & _
				"             product_code," & _
				"             month_saldo_awal," & _
				"             year_saldo_awal," & _
				"             saldo_qty_awal," & _
				"             saldo_amount_awal," & _
				"             Price_saldo_awal," & _
				"             month_prod_sbi," & _
				"             year_prod_sbi," & _
				"             qty_prod_sbi," & _
				"             amount_prod_sbi," & _
				"             Price_prod_sbi," & _
				"             month_qty_sales_sbi," & _
				"             year_qty_sales_sbi," & _
				"             qty_sales_sbi," & _
				"             HPP_Qty_level_1," & _
				"             HPP_Amount_level_1," & _
				"             HPP_Price_level_1," & _
				"             HPP_Qty_level_2," & _
				"             HPP_Amount_level_2," & _
				"             HPP_Price_level_2," & _
				"             (HPP_Amount_level_1+HPP_Amount_level_2) as amount_sales_sbi," & _
				"             decode(qty_sales_sbi,0,0,((HPP_Amount_level_1+HPP_Amount_level_2)/qty_sales_sbi)) as price_sales_sbi," & _
				"             (saldo_qty_awal+qty_prod_sbi-qty_sales_sbi) as saldo_Qty_akhir," & _
				"             (saldo_amount_awal+amount_prod_sbi-HPP_Amount_level_1-HPP_Amount_level_2) as saldo_amount_akhir," & _
				"             case when (saldo_qty_awal+qty_prod_sbi-qty_sales_sbi) <> 0 then ((saldo_amount_awal+amount_prod_sbi-HPP_Amount_level_1-HPP_Amount_level_2)/(saldo_qty_awal+qty_prod_sbi-qty_sales_sbi)) else 0 end price_saldo_akhir" & _
				"             from (" & _
				"                 select" & _
				"                 fcba," & _
				"                 product_code," & _
				"                 month_saldo_awal," & _
				"                 year_saldo_awal," & _
				"                 saldo_qty_awal," & _
				"                 saldo_amount_awal," & _
				"                 Price_saldo_awal," & _
				"                 month_prod_sbi," & _
				"                 year_prod_sbi," & _
				"                 qty_prod_sbi," & _
				"                 amount_prod_sbi," & _
				"                 Price_prod_sbi," & _
				"                 month_qty_sales_sbi," & _
				"                 year_qty_sales_sbi," & _
				"                 qty_sales_sbi," & _
				"                 case when" & _
				"                     qty_sales_sbi >= saldo_qty_awal then saldo_qty_awal" & _
				"                     else qty_sales_sbi" & _
				"                     end as HPP_Qty_level_1," & _
				"                 case when" & _
				"                     qty_sales_sbi >= saldo_qty_awal then Price_saldo_awal" & _
				"                     else Price_saldo_awal" & _
				"                     end as HPP_Price_level_1," & _
				"                 case when" & _
				"                     qty_sales_sbi >= saldo_qty_awal then (saldo_qty_awal*Price_saldo_awal)" & _
				"                     else (saldo_qty_awal*Price_saldo_awal)" & _
				"                     end as HPP_Amount_level_1," & _
				"                 case when" & _
				"                     qty_sales_sbi >= saldo_qty_awal then (qty_sales_sbi-saldo_qty_awal)" & _
				"                     else 0" & _
				"                     end as HPP_Qty_level_2," & _
				"                 case when" & _
				"                     qty_sales_sbi >= saldo_qty_awal then Price_prod_sbi" & _
				"                     else 0" & _
				"                     end as HPP_Price_level_2," & _
				"                 case when" & _
				"                     qty_sales_sbi >= saldo_qty_awal then ((qty_sales_sbi-saldo_qty_awal)*Price_prod_sbi)" & _
				"                     else 0" & _
				"                     end as HPP_Amount_level_2" & _
				"                 from (" & _
				"                     select" & _
				"                     fcba," & _
				"                     product_code," & _
				"                     max(month_saldo_awal) as month_saldo_awal," & _
				"                     max(year_saldo_awal) as year_saldo_awal," & _
				"                     max(saldo_qty_awal) as saldo_qty_awal," & _
				"                     max(saldo_amount_awal) as saldo_amount_awal," & _
				"                     max(Price_saldo_awal) as Price_saldo_awal," & _
				"                     max(month_prod_sbi) as month_prod_sbi," & _
				"                     max(year_prod_sbi) as year_prod_sbi," & _
				"                     max(qty_prod_sbi) as qty_prod_sbi," & _
				"                     max(amount_prod_sbi) as amount_prod_sbi," & _
				"                     max(Price_prod_sbi) as Price_prod_sbi," & _
				"                     max(month_qty_sales_sbi) as month_qty_sales_sbi," & _
				"                     max(year_qty_sales_sbi) as year_qty_sales_sbi," & _
				"                     max(qty_sales_sbi) as qty_sales_sbi" & _
				"                     from (" & _
				"                         select" & _
				"                         fcba," & _
				"                         product_code," & _
				"                         month_saldo_awal," & _
				"                         year_saldo_awal," & _
				"                         qty_cpo_prod_awal as saldo_qty_awal," & _
				"                         prop_amount_cpo_prod_awal as saldo_amount_awal," & _
				"                         case when qty_cpo_prod_awal <> 0 then (prop_amount_cpo_prod_awal/qty_cpo_prod_awal) else 0 end Price_saldo_awal," & _
				"                         0 as month_prod_sbi," & _
				"                         0 as year_prod_sbi," & _
				"                         0 as qty_prod_sbi," & _
				"                         0 as amount_prod_sbi," & _
				"                         0 as Price_prod_sbi," & _
				"                         0 as month_qty_sales_sbi," & _
				"                         0 as year_qty_sales_sbi," & _
				"                         0 as qty_sales_sbi" & _
				"                         from " & _
				"                         (" & _
				"                             select" & _
				"                             fcba," & _
				"                             'CPO' as product_code," & _
				"                             month_saldo_awal," & _
				"                             year_saldo_awal," & _
				"                             qty_cpo_prod_awal," & _
				"                             prop_amount_cpo_prod_awal" & _
				"                             from (" & _
				"                                 select" & _
				"                                 qty_awal.fcba," & _
				"                                 qty_awal.month_saldo_awal," & _
				"                                 qty_awal.year_saldo_awal," & _
				"                                 qty_awal.qty_cpo_prod_awal," & _
				"                                 qty_awal.qty_ker_prod_awal," & _
				"                                 biaya_awal.saldo_amount_awal," & _
				"                                 case when (qty_awal.qty_cpo_prod_awal+qty_awal.qty_ker_prod_awal) <> 0 then (biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2))) else 0 end prop_cpo_prod_awal," & _
				"                                 case when (qty_awal.qty_cpo_prod_awal+qty_awal.qty_ker_prod_awal) <> 0 then ((biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2)))/2) else 0 end prop_ker_prod_awal," & _
				"                                 case when (qty_awal.qty_cpo_prod_awal+qty_awal.qty_ker_prod_awal) <> 0 then round(qty_awal.qty_cpo_prod_awal*((biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2)))),2)"& _
				" 								  else 0 end prop_amount_cpo_prod_awal," & _
				"                                 case when (qty_awal.qty_cpo_prod_awal+qty_awal.qty_ker_prod_awal) <> 0 then round(qty_awal.qty_ker_prod_awal*(((biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2)))/2)),2)" & _
				" 								  else 0 end prop_amount_ker_prod_awal" & _
				"                                 from (" & _
				"                                     select" & _
				"                                     fcba,month_saldo_awal,year_saldo_awal," & _
				"                                     sum(decode(product_code,'CPO',saldo_qty_awal,0)) as qty_cpo_prod_awal," & _
				"                                     sum(decode(product_code,'KER',saldo_qty_awal,0)) as qty_ker_prod_awal" & _
				"                                     from (" & _
				"                                         select" & _
				"                                         iplasprod.mill_summary_tab.fcba," & _
				"                                         case" & _
				"                                             when iplasprod.mill_summary_tab.codedesc = 'CPO003' then 'CPO'" & _
				"                                             else 'KER'" & _
				"                                             end as product_code," & _
				"                                         to_number(to_char(iplasprod.mill_summary_tab.ondate,'mm')) as month_saldo_awal," & _
				"                                         to_number(to_char(iplasprod.mill_summary_tab.ondate,'yyyy')) as year_saldo_awal," & _
				"                                         nvl(iplasprod.mill_summary_tab.todate,0) as saldo_qty_awal " & _
				"                                         from iplasprod.mill_summary_tab " & _
				"                                         where iplasprod.mill_summary_tab.fcba = '"& v_txtfcba &"'" & _
				"                                         and iplasprod.mill_summary_tab.codedesc in ( 'CPO003','PK003')" & _
				"                                         and iplasprod.mill_summary_tab.ondate = to_date('"& v_txtMonth &"/"& v_txtYear &"','mm/yyyy')-1" & _ 
				"                                     )" & _
				"                                     group by fcba,month_saldo_awal,year_saldo_awal" & _
				"                                 ) qty_awal" & _
				"                                 inner join (" & _
				"                                     select iplasprod.cost_activity.fcba as fcba," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'mm'))) as month_saldo_awal," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'yyyy'))) as year_saldo_awal," & _
				"                                     sum(iplasprod.cost_activity.amount) as saldo_amount_awal " & _
				"                                     from iplasprod.cost_activity" & _
				"                                     where cost_activity.fcba = '"& v_txtfcba &"'" & _
				"                                     and to_number(to_char(iplasprod.cost_activity.tdate,'yyyy'))= "& v_txtYear &"-1" & _
				"                                     and cost_activity.jobcode in (" & _
				"                                         select iplasprod.parameterdetail.fccode from iplasprod.parameterdetail where " & _
				"                                         iplasprod.parameterdetail.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                         and iplasprod.parameterdetail.parheadcode = 'CPO'" & _
				"                                         and iplasprod.parameterdetail.fcba = '"& v_txtfcba &"'" & _
				"                                     )" & _
				"                                     group by iplasprod.cost_activity.fcba" & _
				"                                 ) biaya_awal" & _
				"                                 on (" & _
				"                                 qty_awal.fcba = biaya_awal.fcba and" & _
				"                                 qty_awal.month_saldo_awal = biaya_awal.month_saldo_awal and" & _
				"                                 qty_awal.year_saldo_awal = biaya_awal.year_saldo_awal)" & _
				"                             )" & _
				"                             union all" & _
				"                             select" & _
				"                             fcba," & _
				"                             'KER' as product_code," & _
				"                             month_saldo_awal," & _
				"                             year_saldo_awal," & _
				"                             qty_ker_prod_awal," & _
				"                             prop_amount_ker_prod_awal" & _
				"                             from (" & _
				"                                 select" & _
				"                                 qty_awal.fcba," & _
				"                                 qty_awal.month_saldo_awal," & _
				"                                 qty_awal.year_saldo_awal," & _
				"                                 qty_awal.qty_cpo_prod_awal," & _
				"                                 qty_awal.qty_ker_prod_awal," & _
				"                                 biaya_awal.saldo_amount_awal," & _
				"                                 case when (qty_awal.qty_cpo_prod_awal+qty_awal.qty_ker_prod_awal) <> 0 then (biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2))) else 0 end prop_cpo_prod_awal," & _
				"                                 case when (qty_awal.qty_cpo_prod_awal+qty_awal.qty_ker_prod_awal) <> 0 then ((biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2)))/2) else 0 end prop_ker_prod_awal," & _
				"                                 case when (qty_awal.qty_cpo_prod_awal+qty_awal.qty_ker_prod_awal) <> 0 then round(qty_awal.qty_cpo_prod_awal*((biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2)))),2) " & _
				"								  else 0 end prop_amount_cpo_prod_awal," & _
				"                                 case when (qty_awal.qty_cpo_prod_awal+qty_awal.qty_ker_prod_awal) <> 0 then round(qty_awal.qty_ker_prod_awal*(((biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2)))/2)),2)" & _
				" 								  else 0 end prop_amount_ker_prod_awal" & _
				"                                 from (" & _
				"                                     select" & _
				"                                     fcba,month_saldo_awal,year_saldo_awal," & _
				"                                     sum(decode(product_code,'CPO',saldo_qty_awal,0)) as qty_cpo_prod_awal," & _
				"                                     sum(decode(product_code,'KER',saldo_qty_awal,0)) as qty_ker_prod_awal" & _
				"                                     from (" & _
				"                                         select" & _
				"                                         iplasprod.mill_summary_tab.fcba," & _
				"                                         case" & _
				"                                             when iplasprod.mill_summary_tab.codedesc = 'CPO003' then 'CPO'" & _
				"                                             else 'KER'" & _
				"                                             end as product_code," & _
				"                                         to_number(to_char(iplasprod.mill_summary_tab.ondate,'mm')) as month_saldo_awal," & _
				"                                         to_number(to_char(iplasprod.mill_summary_tab.ondate,'yyyy')) as year_saldo_awal," & _
				"                                         nvl(iplasprod.mill_summary_tab.todate,0) as saldo_qty_awal " & _
				"                                         from iplasprod.mill_summary_tab " & _
				"                                         where iplasprod.mill_summary_tab.fcba = '"& v_txtfcba &"'" & _
				"                                         and iplasprod.mill_summary_tab.codedesc in ( 'CPO003','PK003')" & _
				"                                         and iplasprod.mill_summary_tab.ondate = to_date('"& v_txtMonth &"/"& v_txtYear &"','mm/yyyy')-1" & _ 
				"                                     )" & _
				"                                     group by fcba,month_saldo_awal,year_saldo_awal" & _
				"                                 ) qty_awal" & _
				"                                 inner join (" & _
				"                                     select iplasprod.cost_activity.fcba as fcba," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'mm'))) as month_saldo_awal," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'yyyy'))) as year_saldo_awal," & _
				"                                     sum(iplasprod.cost_activity.amount) as saldo_amount_awal " & _
				"                                     from iplasprod.cost_activity" & _
				"                                     where cost_activity.fcba = '"& v_txtfcba &"'" & _
				"                                     and to_number(to_char(iplasprod.cost_activity.tdate,'yyyy'))= "& v_txtYear &"-1" & _
				"                                     and cost_activity.jobcode in (" & _
				"                                         select iplasprod.parameterdetail.fccode from iplasprod.parameterdetail where " & _
				"                                         iplasprod.parameterdetail.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                         and iplasprod.parameterdetail.parheadcode = 'KER'" & _
				"                                         and iplasprod.parameterdetail.fcba = '"& v_txtfcba &"'" & _
				"                                     )" & _
				"                                     group by iplasprod.cost_activity.fcba" & _
				"                                 ) biaya_awal" & _
				"                                 on (" & _
				"                                 qty_awal.fcba = biaya_awal.fcba and" & _
				"                                 qty_awal.month_saldo_awal = biaya_awal.month_saldo_awal and" & _
				"                                 qty_awal.year_saldo_awal = biaya_awal.year_saldo_awal)" & _
				"                             )" & _
				"                         )" & _
				"                         where product_code in ('CPO','KER')" & _
				"                         union all" & _
				"                         select            " & _
				"                         fcba," & _
				"                         product_code," & _
				"                         0 as month_saldo_awal," & _
				"                         0 as year_saldo_awal," & _
				"                         0 as saldo_qty_awal," & _
				"                         0 as saldo_amount_awal," & _
				"                         0 as Price_saldo_awal," & _
				"                         month_prod_sbi," & _
				"                         year_prod_sbi," & _
				"                         qty_cpo_prod_sbi as qty_prod_sbi," & _
				"                         prop_amount_cpo_prod_sbi as amount_prod_sbi," & _
				"                         case when qty_cpo_prod_sbi <> 0 then (prop_amount_cpo_prod_sbi/qty_cpo_prod_sbi) else 0 end Price_prod_sbi," & _
				"                         0 as month_qty_sales_sbi," & _
				"                         0 as year_qty_sales_sbi," & _
				"                         0 as qty_sales_sbi" & _
				"                         from (" & _
				"                             select" & _
				"                             fcba," & _
				"                             'CPO' as product_code," & _
				"                             month_prod_sbi," & _
				"                             year_prod_sbi," & _
				"                             qty_cpo_prod_sbi," & _
				"                             prop_amount_cpo_prod_sbi" & _
				"                             from (" & _
				"                                 select" & _
				"                                 qty_sbi.fcba," & _
				"                                 qty_sbi.month_prod_sbi," & _
				"                                 qty_sbi.year_prod_sbi," & _
				"                                 qty_sbi.qty_cpo_prod_sbi," & _
				"                                 qty_sbi.qty_ker_prod_sbi," & _
				"                                 biaya_sbi.amount_prod_sbi," & _
				"                                 case when (qty_sbi.qty_cpo_prod_sbi+qty_sbi.qty_ker_prod_sbi) <> 0 then (biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2))) else 0 end prop_cpo_prod_sbi," & _
				"                                 case when (qty_sbi.qty_cpo_prod_sbi+qty_sbi.qty_ker_prod_sbi) <> 0 then ((biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2)))/2) else 0 end prop_ker_prod_sbi," & _
				"                                 case when (qty_sbi.qty_cpo_prod_sbi+qty_sbi.qty_ker_prod_sbi) <> 0 then round(qty_sbi.qty_cpo_prod_sbi*((biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2)))),2) "  & _
				"								  else 0 end prop_amount_cpo_prod_sbi," & _
				"                                 case when (qty_sbi.qty_cpo_prod_sbi+qty_sbi.qty_ker_prod_sbi) <> 0 then round(qty_sbi.qty_ker_prod_sbi*(((biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2)))/2)),2)" & _
				" 								  else 0 end prop_amount_ker_prod_sbi" & _
				"                                 from (" & _
				"                                     select" & _
				"                                     fcba,month_prod_sbi,year_prod_sbi," & _
				"                                     sum(decode(product_code,'CPO',qty_prod_sbi,0)) as qty_cpo_prod_sbi," & _
				"                                     sum(decode(product_code,'KER',qty_prod_sbi,0)) as qty_ker_prod_sbi" & _
				"                                     from (" & _
				"                                         select iplasprod.mill_production.fcba as fcba,iplasprod.mill_production.product_code," & _
				"                                         max(to_number(to_char(iplasprod.mill_production.tdate,'mm'))) as month_prod_sbi," & _
				"                                         max(to_number(to_char(iplasprod.mill_production.tdate,'yyyy'))) as year_prod_sbi," & _
				"                                         sum(iplasprod.mill_production.production_today) as qty_prod_sbi " & _
				"                                         from iplasprod.mill_production" & _
				"                                         where iplasprod.mill_production.fcba = '"& v_txtfcba &"'" & _
				"                                         and (iplasprod.mill_production.tdate between to_date(to_char(trunc(sysdate,'year'),'mm/dd/yyyy'),'mm/dd/yyyy') " & _
				"										  and to_date('"& v_txtdate &"','mm/dd/yyyy'))" & _
				"                                         and iplasprod.mill_production.product_code in ('CPO','KER')" & _
				"                                         group by iplasprod.mill_production.fcba,iplasprod.mill_production.product_code" & _
				"                                     )" & _
				"                                     group by fcba,month_prod_sbi,year_prod_sbi" & _
				"                                 ) qty_sbi" & _
				"                                 inner join (" & _
				"                                     select iplasprod.cost_activity.fcba as fcba," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'mm'))) as month_prod_sbi," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'yyyy'))) as year_prod_sbi," & _
				"                                     sum(iplasprod.cost_activity.amount) as amount_prod_sbi " & _
				"                                     from iplasprod.cost_activity" & _
				"                                     where cost_activity.fcba = '"& v_txtfcba &"'" & _
				"                                     and (iplasprod.cost_activity.tdate between to_date(to_char(trunc(sysdate,'year'),'mm/dd/yyyy'),'mm/dd/yyyy') " & _
				"									  and to_date('"& v_txtdate &"','mm/dd/yyyy'))" & _
				"                                     and cost_activity.jobcode in (" & _
				"                                         select iplasprod.parameterdetail.fccode from iplasprod.parameterdetail where " & _
				"                                         iplasprod.parameterdetail.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                         and iplasprod.parameterdetail.parheadcode = 'CPO'" & _
				"                                         and iplasprod.parameterdetail.fcba = '"& v_txtfcba &"'" & _
				"                                     )" & _
				"                                     group by iplasprod.cost_activity.fcba" & _
				"                                 ) biaya_sbi" & _
				"                                 on (" & _
				"                                 qty_sbi.fcba = biaya_sbi.fcba and" & _
				"                                 qty_sbi.month_prod_sbi = biaya_sbi.month_prod_sbi and" & _
				"                                 qty_sbi.year_prod_sbi = biaya_sbi.year_prod_sbi)" & _
				"                             )" & _
				"                             union all" & _
				"                             select" & _
				"                             fcba," & _
				"                             'KER' as product_code," & _
				"                             month_prod_sbi," & _
				"                             year_prod_sbi," & _
				"                             qty_ker_prod_sbi," & _
				"                             prop_amount_ker_prod_sbi" & _
				"                             from (" & _
				"                                 select" & _
				"                                 qty_sbi.fcba," & _
				"                                 qty_sbi.month_prod_sbi," & _
				"                                 qty_sbi.year_prod_sbi," & _
				"                                 qty_sbi.qty_cpo_prod_sbi," & _
				"                                 qty_sbi.qty_ker_prod_sbi," & _
				"                                 biaya_sbi.amount_prod_sbi," & _
				"                                 case when (qty_sbi.qty_cpo_prod_sbi+qty_sbi.qty_ker_prod_sbi) <> 0 then  (biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2))) else 0 end prop_cpo_prod_sbi," & _
				"                                 case when (qty_sbi.qty_cpo_prod_sbi+qty_sbi.qty_ker_prod_sbi) <> 0 then ((biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2)))/2) else 0 end prop_ker_prod_sbi," & _
				"                                 case when (qty_sbi.qty_cpo_prod_sbi+qty_sbi.qty_ker_prod_sbi) <> 0 then  round(qty_sbi.qty_cpo_prod_sbi*((biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2)))),2) " & _
				"								  else 0 end prop_amount_cpo_prod_sbi," & _
				"                                 case when (qty_sbi.qty_cpo_prod_sbi+qty_sbi.qty_ker_prod_sbi) <> 0 then round(qty_sbi.qty_ker_prod_sbi*(((biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2)))/2)),2)" & _
				" 								  else 0 end prop_amount_ker_prod_sbi" & _
				"                                 from (" & _
				"                                     select" & _
				"                                     fcba,month_prod_sbi,year_prod_sbi," & _
				"                                     sum(decode(product_code,'CPO',qty_prod_sbi,0)) as qty_cpo_prod_sbi," & _
				"                                     sum(decode(product_code,'KER',qty_prod_sbi,0)) as qty_ker_prod_sbi" & _
				"                                     from (" & _
				"                                         select iplasprod.mill_production.fcba as fcba,iplasprod.mill_production.product_code," & _
				"                                         max(to_number(to_char(iplasprod.mill_production.tdate,'mm'))) as month_prod_sbi," & _
				"                                         max(to_number(to_char(iplasprod.mill_production.tdate,'yyyy'))) as year_prod_sbi," & _
				"                                         sum(iplasprod.mill_production.production_today) as qty_prod_sbi " & _
				"                                         from iplasprod.mill_production" & _
				"                                         where iplasprod.mill_production.fcba = '"& v_txtfcba &"'" & _
				"                                         and (iplasprod.mill_production.tdate between to_date(to_char(trunc(sysdate,'year'),'mm/dd/yyyy'),'mm/dd/yyyy') " & _
				"										  and to_date('"& v_txtdate &"','mm/dd/yyyy'))" & _
				"                                         and iplasprod.mill_production.product_code in ('CPO','KER')" & _
				"                                         group by iplasprod.mill_production.fcba,iplasprod.mill_production.product_code" & _
				"                                     )" & _
				"                                     group by fcba,month_prod_sbi,year_prod_sbi" & _
				"                                 ) qty_sbi" & _
				"                                 inner join (" & _
				"                                     select iplasprod.cost_activity.fcba as fcba," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'mm'))) as month_prod_sbi," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'yyyy'))) as year_prod_sbi," & _
				"                                     sum(iplasprod.cost_activity.amount) as amount_prod_sbi " & _
				"                                     from iplasprod.cost_activity" & _
				"                                     where cost_activity.fcba = '"& v_txtfcba &"'" & _
				"                                     and (iplasprod.cost_activity.tdate between to_date(to_char(trunc(sysdate,'year'),'mm/dd/yyyy'),'mm/dd/yyyy') " & _
				"									  and to_date('"& v_txtdate &"','mm/dd/yyyy'))" & _
				"                                     and cost_activity.jobcode in (" & _
				"                                         select iplasprod.parameterdetail.fccode from iplasprod.parameterdetail where " & _
				"                                         iplasprod.parameterdetail.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                         and iplasprod.parameterdetail.parheadcode = 'KER'" & _
				"                                         and iplasprod.parameterdetail.fcba = '"& v_txtfcba &"'" & _
				"                                     )" & _
				"                                     group by iplasprod.cost_activity.fcba" & _
				"                                 ) biaya_sbi" & _
				"                                 on (" & _
				"                                 qty_sbi.fcba = biaya_sbi.fcba and" & _
				"                                 qty_sbi.month_prod_sbi = biaya_sbi.month_prod_sbi and" & _
				"                                 qty_sbi.year_prod_sbi = biaya_sbi.year_prod_sbi)" & _
				"                             )" & _
				"                         )" & _
				"                         where product_code in ('CPO','KER')" & _
				"                         union all" & _
				"                         select" & _
				"                         fcba," & _
				"                         product_code," & _
				"                         0 as month_saldo_awal," & _
				"                         0 as year_saldo_awal," & _
				"                         0 as saldo_qty_awal," & _
				"                         0 as saldo_amount_awal," & _
				"                         0 as Price_saldo_awal," & _
				"                         0 as month_prod_sbi," & _
				"                         0 as year_prod_sbi," & _
				"                         0 as qty_prod_sbi," & _
				"                         0 as amount_prod_sbi," & _
				"                         0 as Price_prod_sbi," & _
				"                         month_qty_sales_sbi," & _
				"                         year_qty_sales_sbi," & _
				"                         qty_sales_sbi" & _
				"                         from (" & _
				"                             select " & _
				"                             fcba," & _
				"                             product_code," & _
				"                             month_qty_sales_sbi," & _
				"                             year_qty_sales_sbi," & _
				"                             weightin," & _
				"                             weightout," & _
				"                             (weightout-weightin) as qty_sales_sbi" & _
				"                             from (        " & _
				"                                 select iplasprod.weighbridgeticket.fcba,iplasprod.weighbridgeticket.product_productcode as product_code," & _
				"                                 max(to_number(to_char(iplasprod.weighbridgeticket.dateout,'mm'))) as month_qty_sales_sbi," & _
				"                                 max(to_number(to_char(iplasprod.weighbridgeticket.dateout,'yyyy'))) as year_qty_sales_sbi," & _
				"                                 sum(iplasprod.weighbridgeticket.weightin) as weightin," & _
				"                                 sum(iplasprod.weighbridgeticket.weightout) as weightout" & _
				"                                 from iplasprod.weighbridgeticket" & _
				"                                 where iplasprod.weighbridgeticket.fcba = '"& v_txtfcba &"'" & _
				"                                 and (iplasprod.weighbridgeticket.dateout between to_date(to_char(trunc(sysdate,'year'),'mm/dd/yyyy'),'mm/dd/yyyy')" & _
				" 								  and to_date('"& v_txtdate &"','mm/dd/yyyy'))" & _
				"                                 and iplasprod.weighbridgeticket.product_productcode in ('CPO','KER')" & _
				"                                 group by iplasprod.weighbridgeticket.fcba,iplasprod.weighbridgeticket.product_productcode" & _
				"                             )" & _
				"                         )  " & _
				"                     ) " & _
				"                     group by fcba,product_code" & _
				"                 )" & _
				"             )" & _
				"         )" & _
				"     )" & _
				" )" 

' response.write strPusingan
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_field = qORS.getrows()
	arr_field_eof = false
End If
qORS.Close
set qORS = Nothing

%>

<%	
    strTitle = 	" select PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPPLANTATION') titleGP, " & _
			    " PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','PLANTATION') titleP , " & _
				" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPREGION') titleGR, " & _
				" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','REGION') titleR, " & _
				" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','COMPANY') titleC, " & _
				" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','SITE') titleS from dual"
						
	Set rsTitle = Server.CreateObject("ADODB.Recordset")
	'response.write rsTitle
	rsTitle.Open strTitle, connstr, 3,3	

	if not rsTitle.EOF then
		no = 0
		titleGroupPlantation = rsTitle ("titleGP")			
		titlePlantation = rsTitle ("titleP")
		titleGroupRegion = rsTitle ("titleGR")
		titleRegion = rsTitle ("titleR")
		titleCompany = rsTitle ("titleC")
		titleSite = rsTitle ("titleS")
	else
		titleGroupPlantation = "Group Plantation"			
		titlePlantation = "Plantation"
		titleGroupRegion = "Group Region"
		titleRegion = "Region"
		titleCompany = "Company"
		titleSite = "Site"
	end if
	rsTitle.Close
	set rsTitle = Nothing
			
%>
<html>

<head>
<title>ePlaS Report - Sales Harga Pokok Penjualan</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
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
<script>
 $(function() {
	$( ".datepicker" ).datepicker({
		changeMonth: true,
		changeYear: true
	});
});
</script>
<!--
<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
-->

<style type="text/css">
@media print
{
.printhide{display:none;}
}

th.rotate-270 {
  height: 200px; /* 1 */
  white-space: nowrap;
  min-width:40px;
  max-width:40px;
  overflow:hidden;
}
th.rotate-270 > div {
	transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-webkit-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-ms-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
    width:15px; /* 4 */
	font-size:10px;
	font-weight:normal;
}

/*
1. Mengatur tinggi table head.
2. mengatur posisi X div
3. mengatur posisi Y div.
4. mengatur lebarnya DIV yang akan mempengaruhi Lebarnya Table Head.
*/
</style>


<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />
</head>

<body>
	<div id="wrapper">
	 <!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
				<div class="printhide">
					<table id="form-table" class="centered-margin">
						<tr>
							<th>PABRIK</th>
							<th>TAHUN</th>
							<th>BULAN</th>
							<th>TANGGAL</th>
						</tr>
						<tr>
							<td>
								<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
									<%
									Dim rsFcba
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.CursorLocation = adUseClient
									strsql =    " SELECT '--ALL--' as fcba from DUAL UNION " & _  
									            " select fccode fcba from businessunit " & _
												" where " & _
												" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL AND fcType = 'M'" 
									  
									
									rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsFcba.eof then
										DO while not rsFcba.eof
											if v_txtFcba = rsFcba("fcba") then
											%>			
												<option value="<%=rsFcba("fcba")%>" Selected><%=rsFcba("fcba")%></option>
											<% Else %>	
												<option value="<%=rsFcba("fcba")%>"><%=rsFcba("fcba")%></option>
											<% End If
											
											rsFcba.movenext
										loop
									End if
									rsFcba.Close
									set rsFcba = Nothing
									%>
								</select>
							</td> 
								<% 
									Dim mon, Yr, MinYr, MaxYr, i, MM
									mon = Month(Session("v_Costbook"))
									If len(v_txtMonth)=0 then v_txtMonth=mon

									Yr = Year(Session("v_Costbook"))
									MinYr = 2015
									MaxYr = Yr  + 2
									If len(v_txtYear)= 0 then v_txtYear = Yr
								%>
							<td width="4">
								<select name="txtYear">
									<% For i = MinYr to MaxYr %>
										<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
									<% Next %>
								</select>
							</td> 
							<td width="2">
								<select name="txtMonth">
								<% 
									For i = 1 to 12		
										if i <= 9 then 
											iBulan = 0&i
										else
											iBulan = i
										end if
										i	= iBulan
								%>
									<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
								<%
									Next
								%>
								</select>
							</td> 
							<td width="2">
								<input class="datepicker" type="text" name="txtDate" value="<%=v_txtDate%>" />
							</td>
							<td>
								<input type="submit" value="View" />
							</td>
						</tr>				
					</table>
				</div>
			</form>
		</div>
<table border="0" width="100%" id="product-table" class="xthf-site" >
	<thead style="border-left:1px #000000 solid;">
		<tr>
	         <th colspan="<%=d_span%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
			 <div id="header-wrapper">
				
				<div id="subheader" class="centered">
					<h2>LAPORAN SALES HARGA POKOK PENJUALAN</h2>
					<br/>
					<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
				</div>
	          </div>  
			 </th>
		</tr>
	 	<tr>
			<th rowspan="2">KEBUN</th>
			<th rowspan="2">JENIS</th>
			<th colspan="3">SALDO AWAL</th>
			<th colspan="3">PENERIMAAN</th>
			<th colspan="3">PENGELUARAN</th>
			<th colspan="4">LAIN - LAIN</th>
			<th colspan="3">SALDO AKHIR</th>
		</tr>
		<tr>
			<th>QTY</th>
			<th>@</th>
			<th>JUMLAH</th>

			<th>QTY</th>
			<th>@</th>
			<th>JUMLAH</th>

			<th>QTY</th>
			<th>@</th>
			<th>JUMLAH</th>

			<th>QTY</th>
			<th>@</th>
			<th>JUMLAH</th>
			<th>KETERANGAN</th>

			<th>QTY</th>
			<th>@</th>
			<th>JUMLAH</th>
		</tr>
    </thead>
	<tbody>
		<%
		Set rsCat = connStr.Execute(strPusingan)
		v_saldo_qty_awal = 0
		v_saldo_amount_awal = 0 
		v_qty_prod_sbi = 0 
		v_amount_prod_sbi = 0
		v_qty_sales_sbi = 0
		v_amount_sales_sbi= 0
		v_saldo_Qty_akhir = 0
		v_saldo_amount_akhir = 0

		if Not rsCat.eof then
			While Not rsCat.EOF
			' TOTALBULANAN = cdbl(to_nol(rsCat("JANUARI_TTL"))) + cdbl(to_nol(rsCat("FEBRUARI_TTL"))) + cdbl(to_nol(rsCat("MARET_TTL"))) + cdbl(to_nol(rsCat("APRIL_TTL"))) + cdbl(to_nol(rsCat("MEI_TTL"))) + cdbl(to_nol(rsCat("JUNI_TTL"))) + cdbl(to_nol(rsCat("JULI_TTL"))) + cdbl(to_nol(rsCat("AGUSTUS_TTL"))) + cdbl(to_nol(rsCat("SEPTEMBER_TTL"))) + cdbl(to_nol(rsCat("OKTOBER_TTL"))) + cdbl(to_nol(rsCat("NOVEMBER_TTL"))) + cdbl(to_nol(rsCat("DESEMBER_TTL")))

			' TOTALHARIAN = cdbl(to_nol(rsCat("JANUARI"))) + cdbl(to_nol(rsCat("FEBRUARI"))) + cdbl(to_nol(rsCat("MARET"))) + cdbl(to_nol(rsCat("APRIL"))) + cdbl(to_nol(rsCat("MEI"))) + cdbl(to_nol(rsCat("JUNI"))) + cdbl(to_nol(rsCat("JULI"))) + cdbl(to_nol(rsCat("AGUSTUS"))) + cdbl(to_nol(rsCat("SEPTEMBER"))) + cdbl(to_nol(rsCat("OKTOBER"))) + cdbl(to_nol(rsCat("NOVEMBER"))) + cdbl(to_nol(rsCat("DESEMBER"))) 
		%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("product_code")%></td> 
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("saldo_qty_awal"))),0)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("Price_saldo_awal"))),0)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("saldo_amount_awal"))),0)%></td>

			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("qty_prod_sbi"))),0)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("Price_prod_sbi"))),0)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("amount_prod_sbi"))),0)%></td>

			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("qty_sales_sbi"))),0)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("price_sales_sbi"))),0)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("amount_sales_sbi"))),0)%></td> 

			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"></td> 

			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("saldo_Qty_akhir"))),0)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("price_saldo_akhir"))),0)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("saldo_amount_akhir"))),0)%></td> 
		</tr>
	<%
		v_saldo_qty_awal = v_saldo_qty_awal +  cdbl(to_nol(rscat("saldo_qty_awal"))) 
		v_saldo_amount_awal = v_saldo_amount_awal +  cdbl(to_nol(rscat("saldo_amount_awal")))

		v_qty_prod_sbi = v_qty_prod_sbi +  cdbl(to_nol(rscat("qty_prod_sbi"))) 
		v_amount_prod_sbi = v_amount_prod_sbi +  cdbl(to_nol(rscat("amount_prod_sbi")))

		v_qty_sales_sbi = v_qty_sales_sbi +  cdbl(to_nol(rscat("qty_sales_sbi"))) 
		v_amount_sales_sbi = v_amount_sales_sbi +  cdbl(to_nol(rscat("amount_sales_sbi")))

		v_saldo_Qty_akhir = v_saldo_Qty_akhir +  cdbl(to_nol(rscat("saldo_Qty_akhir"))) 
		v_saldo_amount_akhir = v_saldo_amount_akhir +  cdbl(to_nol(rscat("saldo_amount_akhir")))
		rsCat.MoveNext
		Wend
	Else
		response.write "<tr><td colspan='18'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
		<tr bgcolor="#F4F4A0">
			<td align="center" colspan="2" style="<%=cmdstyle%>"><font face="Verdana" size="1" style="font-weight: bold;">TOTAL</font></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1" style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_saldo_qty_awal)),0)%></font></td>
			<td align="right" style="<%=cmdstyle%>"></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1" style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_saldo_amount_awal)),0)%></font></td>

			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1" style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_qty_prod_sbi)),0)%></font></td>
			<td align="right" style="<%=cmdstyle%>"></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1" style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_amount_prod_sbi)),0)%></font></td>

			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1" style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_qty_sales_sbi)),0)%></font></td>
			<td align="right" style="<%=cmdstyle%>"></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1" style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_amount_sales_sbi)),0)%></font></td>

			<td align="right" style="<%=cmdstyle%>"></td>
			<td align="right" style="<%=cmdstyle%>"></td>
			<td align="right" style="<%=cmdstyle%>"></td>
			<td align="right" style="<%=cmdstyle%>"></td>

			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1" style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_saldo_Qty_akhir)),0)%></font></td>
			<td align="right" style="<%=cmdstyle%>"></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1" style="font-weight: bold;"><%=formatnumber(cdbl(to_nol(v_saldo_amount_akhir)),0)%></font></td>
		</tr>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</tbody>
</table>


<script src="/plugins/chosen/chosen.jquery.min.js"></script>
	<script type="text/javascript">	
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
		
	</script>
	</body>
</html>