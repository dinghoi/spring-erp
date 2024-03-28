<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
On Error Resume Next
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn=Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder=New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim uploadForm,sales_month,objFile
Dim cn,rs,rowcount,xgr,fldcount,tot_cnt,read_cnt,write_cnt
Dim from_date, end_date, to_date

'���ε� ������Ʈ ��ü ����
Set uploadForm=Server.CreateObject("ABCUpload4.XForm")

uploadForm.AbsolutePath=True
uploadForm.Overwrite=True
uploadForm.MaxUploadSize=1024*1024*50

sales_month = uploadForm("sales_month")
objFile = uploadForm("objFile")

from_date = Mid(sales_month, 1, 4)&"-"&Mid(sales_month, 5, 2)&"-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

Set cn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.Recordset")

cn.Open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ="&objFile&";"
rs.Open "select * from [6:10000]",cn,"0"

rowcount = -1
xgr=rs.getrows
rowcount=UBound(xgr,2)
fldcount=rs.fields.count

tot_cnt = rowcount + 1
read_cnt = 0
write_cnt = 0

'DB Ʈ����� ����
DBConn.BeginTrans

Dim i, sales_date, approve_no, sales_company, trade_no
Dim trade_company, trade_owner, price, cost, cost_vat, imsi_sales_memo, sales_memo
Dim emp_name, rs_trade, trade_name, group_name, rs_emp, rsSales, sales_saupbu
Dim field_check, field_view, url, err_msg

'Dim slip_no, collect_due_date

If rowcount > -1 Then
	For i = 0 To rowcount
		'���� ��ȣ üũ
		If f_toString(xgr(1,i),"") = "" Then
			Exit For
		End If

		sales_date = xgr(0,i)'�ۼ�����
		approve_no = xgr(1,i)'���ι�ȣ
		sales_company = f_SalesCompany(xgr(6,i))'������ ��ȣ
		trade_no = xgr(9,i)'���޹޴��ڻ���ڵ�Ϲ�ȣ
		trade_company = xgr(11,i)'��ȣ(�ŷ�ó)
		trade_owner = xgr(12,i)'��ǥ�ڸ�
		price = CDbl(f_toString(xgr(14,i),0))'�հ�ݾ�
		cost = CDbl(f_toString(xgr(15,i),0))'���ް���
		cost_vat = CDbl(f_toString(xgr(16,i),0))'����

		imsi_sales_memo=xgr(26,i)'ǰ���
		sales_memo=Replace(imsi_sales_memo,"'","")

		emp_name=xgr(33,i)'�����
		saupbu=xgr(34,i)'�μ�

		'��ǥ��ȣ
		'If f_toString(xgr(35,i),"")="" Then
		'	slip_no=""
		'Else
		'	slip_no=Replace(xgr(13,i),",","")
		'End If

		'collect_due_date=xgr(36,i)	'�Աݿ�����

		'If collect_due_date="" Or IsNull(collect_due_date) Then
		'	collect_due_date=""
		'Else
		'	collect_due_date="20"&Replace(collect_due_date,".","-")
		'End If

		trade_no=Replace(trade_no,"-","")

		'�ŷ�ó�� ��ȸ
		objBuilder.Append "SELECT trade_name, group_name FROM trade "
		objBuilder.Append "WHERE trade_no='"&trade_no&"';"

		Set rs_trade=DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_trade.EOF Or rs_trade.BOF Then
			trade_name=trade_company
			group_name=""
		Else
			trade_name=rs_trade("trade_name")
			group_name=rs_trade("group_name")
		End If
		rs_trade.close()

		'��� ��ȸ
		objBuilder.Append "SELECT emp_no FROM emp_master AS emmt "

		If saupbu="��Ÿ�����" Then
			'SQL = "SELECT emp_no FROM emp_master WHERE emp_name = '"&emp_name&"' "
			objBuilder.Append "WHERE emp_name='"&emp_name&"';"
		Else
			'SQL = "SELECT emp_no FROM emp_master AS emmt "
			'SQL = SQL & "INNER JOIN emp_org_mst AS eomt ON eomt.org_code = emmt.emp_org_code "
			'SQL = SQL & "WHERE emmt.emp_name = '"&emp_name&"' AND eomt.org_bonbu = '"&saupbu&"' "
			objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON eomt.org_code=emmt.emp_org_code "
			objBuilder.Append "WHERE emmt.emp_name='"&emp_name&"' AND eomt.org_bonbu='"&saupbu&"';"
		End If

		Set rs_emp = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_emp.EOF Or rs_emp.BOF Then
			emp_no="error"
		Else
			emp_no=rs_emp("emp_no")
		End If
		rs_emp.Close()

		read_cnt = read_cnt + 1'���� ����

		objBuilder.Append "SELECT approve_no FROM saupbu_sales "
		objBuilder.Append "WHERE approve_no='"&approve_no&"';"

		Set rsSales=DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsSales.EOF Or rsSales.BOF Then
			'������ �Է� ó��
			objBuilder.Append "INSERT INTO saupbu_sales(sales_date,approve_no,slip_no,sales_company,saupbu,"
			objBuilder.Append "company,group_name,trade_no,sales_amt,cost_amt,vat_amt,collect_due_date,"
			objBuilder.Append "emp_no,emp_name,sales_memo,reg_id,reg_name,reg_date)VALUES("
			'objBuilder.Append "'"&sales_date&"','"&approve_no&"','"&slip_no&"','"&sales_company&"','"&saupbu&"',"
			objBuilder.Append "'"&sales_date&"','"&approve_no&"','NULL','"&sales_company&"','"&saupbu&"',"
			objBuilder.Append "'"&trade_name&"','"&group_name&"','"&trade_no&"',"&price&","&cost&","&cost_vat&","
			'If IsDate(collect_due_date) Then
			'	objBuilder.Append "'"&collect_due_date&"',"
			'Else
				objBuilder.Append "NULL,"
			'End If
			objBuilder.Append "'"&emp_no&"','"&emp_name&"','"&sales_memo&"','"&user_id&"','"&user_name&"',NOW());"
		Else
			objBuilder.Append "UPDATE saupbu_sales SET "
			objBuilder.Append "	sales_date='"&sales_date&"',slip_no='NULL',sales_company='"&sales_company&"',"
			objBuilder.Append "	saupbu='"&saupbu&"',company='"&trade_name&"',group_name='"&group_name&"',"
			objBuilder.Append "	trade_no='"&trade_no&"',sales_amt="&price&",cost_amt="&cost&",vat_amt="&cost_vat&","
			'If IsDate(collect_due_date) Then
			'	objBuilder.Append "	collect_due_date='"&collect_due_date&"',"
			'Else
				objBuilder.Append "	collect_due_date=NULL,"
			'End If
			objBuilder.Append "	emp_no='"&emp_no&"',emp_name='"&emp_name&"',sales_memo='"&sales_memo&"',"
			objbuilder.Append "	mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=NOW() "
			objbuilder.Append "WHERE approve_no='"&approve_no&"';"
		End If

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
		rsSales.Close()

		write_cnt = write_cnt+1'ó�� ����
	Next
End If
Set rs_trade = Nothing
Set rs_emp = Nothing
Set rsSales = Nothing

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg="ó���� Error�� �߻��Ͽ����ϴ�."
	url = "/finance/sales_bill_upload.asp"
Else
	DBConn.CommitTrans
	err_msg = "�� "&CStr(read_cnt)&"�� �а� "&CStr(write_cnt)&" �� ó���Ǿ����ϴ�."
	url = "/finance/sales_bill_mg.asp?from_date="&from_date&"&to_date="&to_date&"&sales_saupbu=��ü&field_check=total&ck_sw=y"
End If

rs.Close() : Set rs = Nothing
cn.Close() : Set cn = Nothing
DBConn.Close() : Set DBConn = Nothing

'sales_saupbu="��ü"
'field_check="total"
'field_view=""
'ck_sw="y"

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&err_msg&"');"
Response.Write "	location.replace('"&url&"');"
Response.Write "</script>"
Response.End
%>