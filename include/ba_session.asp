<%
' =====================================================================
' Isi ulang session yang bergantung pada Business Unit (BA).
' Dipakai bersama oleh UsersChangedBusinessUnit.asp dan ba_switch.asp
' supaya ganti BA tanpa logout hasilnya sama dengan login ke BA tersebut.
' rsBU: recordset BusinessUnit dengan kolom fccode, fccompanycode, FCNAME,
'       FCCOMPANYNAME, FCMANAGER, FCKTU, FCCROP, FCTYPE.
' cmdFindPeriod / cmdPeriodClose (txtFunction.inc) membaca Session("fcBa"),
' jadi Session("fcBa") wajib diisi lebih dulu.
' =====================================================================
Sub FbpSetSessionBA(rsBU)
	Session("fcBa") = rsBU("fcCode")
	Session("BA") = rsBU("fcName")
	Session("COMPANYCODE") = rsBU("fccompanycode")
	Session("COMPANY") = rsBU("FCCOMPANYNAME")
	Session("FCMANAGER") = rsBU("FCMANAGER")
	Session("KTU") = rsBU("FCKTU")
	Session("Crop") = rsBU("FCCROP")
	Session("Estate_Mill") = rsBU("fcType")

	Session("v_Bank") = cmdFindPeriod("Bank")
	Session("v_Purchase") = cmdFindPeriod("Purchase")
	Session("v_Stores") = cmdFindPeriod("Stores")
	Session("v_Employee") = cmdFindPeriod("Employee")
	Session("v_Workshop") = cmdFindPeriod("Workshop")
	Session("v_Vehicle") = cmdFindPeriod("Vehicle")
	Session("v_Contract") = cmdFindPeriod("Contract")
	Session("v_Nursery") = cmdFindPeriod("Nursery")
	Session("v_Project") = cmdFindPeriod("Project")
	Session("v_Building") = cmdFindPeriod("Building")
	Session("v_General_Charges") = cmdFindPeriod("General_Charges")
	Session("v_Reversed") = cmdFindPeriod("Reversed")
	Session("v_Costbook") = cmdFindPeriod("Costbook")
	Session("v_Harvesting") = cmdFindPeriod("Harvesting")
	Session("v_WeighBridge") = cmdFindPeriod("WeighBridge")
	Session("v_Mill_Production") = cmdFindPeriod("Mill_Production")
	Session("v_Delivery") = cmdFindPeriod("Delivery")
	Session("v_ERPReport") = cmdFindPeriod("ERPReport")

	Session("cls_Bank") = cmdPeriodClose("Bank")
	Session("cls_Purchase") = cmdPeriodClose("Purchase")
	Session("cls_Stores") = cmdPeriodClose("Stores")
	Session("cls_Employee") = cmdPeriodClose("Employee")
	Session("cls_Workshop") = cmdPeriodClose("Workshop")
	Session("cls_Vehicle") = cmdPeriodClose("Vehicle")
	Session("cls_Contract") = cmdPeriodClose("Contract")
	Session("cls_Nursery") = cmdPeriodClose("Nursery")
	Session("cls_Project") = cmdPeriodClose("Project")
	Session("cls_Building") = cmdPeriodClose("Building")
	Session("cls_General_Charges") = cmdPeriodClose("General_Charges")
	Session("cls_Reversed") = cmdPeriodClose("Reversed")
	Session("cls_Costbook") = cmdPeriodClose("Costbook")
	Session("cls_Harvesting") = cmdPeriodClose("Harvesting")
	Session("cls_WeighBridge") = cmdPeriodClose("WeighBridge")
	Session("cls_Mill_Production") = cmdPeriodClose("Mill_Production")
	Session("cls_Delivery") = cmdPeriodClose("Delivery")
	Session("cls_ERPReport") = cmdPeriodClose("ERPReport")

	Session("NextMonth") = cmdPeriodClose("Financial")
End Sub
%>
