<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'On Error Resume Next
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
Dim uploadForm, bill_month, objFile
Dim cn, rs, rowcount, xgr, fldcount, tot_cnt, read_cnt, write_cnt

Dim rs_trade, rsTax, i, bill_date, approve_no, trade_no, trade_name
Dim trade_owner, owner_trade_no, price, cost, cost_vat, bill_collect
Dim send_email, receive_email, imsi_bill_memo, tax_bill_memo
Dim owner_company, err_msg, url

'업로드 컴포넌트 객체 생성
Set uploadForm = Server.CreateObject("ABCUpload4.XForm")
uploadForm.AbsolutePath = True
uploadForm.Overwrite = true
uploadForm.MaxUploadSize = 1024*1024*50

bill_month = uploadForm("bill_month")
objFile = uploadForm("objFile")

Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
rs.Open "select * from [3:10000]",cn,"0"

rowcount=-1
xgr = rs.getrows
rowcount = UBound(xgr,2)
fldcount = rs.fields.count

tot_cnt = rowcount + 1
read_cnt = 0
write_cnt = 0

'DB 트랜잭션 시작
DBConn.BeginTrans

Dim t_bill_date, t_approve, t_owner_company, t_trade_name, t_emp_name, t_emp_no, t_price
Dim t_cost, t_cost_vat, t_tax_bill_memo, t_org_code, t_company, t_mg_saupbu, t_slip_gubun
Dim t_account_str, slip_date, customer, customer_no, pay_method, vat_yn, pay_yn
Dim rsEmp, emp_grade, emp_name, rsGe, max_seq, slip_seq, cost_reg_yn, rsOrg, arr_str, arr_account
Dim t_account, t_account_item, j, slip_gubun, slip_account, rsGeneral, k

If rowcount > -1 Then
	For i = 0 To rowcount
		t_bill_date = f_toString(xgr(0,i), "")'발행일자
		t_approve = f_toString(xgr(1,i), "")'승인번호
		t_owner_company = f_toString(xgr(2,i), "")'계산서소유회사
		t_trade_name = f_toString(xgr(3,i), "")'상호명
		't_price = f_toString(xgr(4,i), 0)'합계
		't_cost = f_toString(xgr(5,i), 0)'공급가액
		't_cost_vat = f_toString(xgr(6,i), 0)'부가세
		t_tax_bill_memo = f_toString(xgr(7,i), "")'거래내역

		't_emp_no = f_toString(xgr(8,i), "")'담당사사번
		t_emp_name = f_toString(xgr(9,i), "")'담당자

		t_org_code = f_toString(xgr(10,i), "")'사용조직코드
		t_company = f_toString(xgr(11,i), "")'고객사
		t_mg_saupbu = f_toString(xgr(12,i), "")'담당사업부
		't_slip_gubun = f_toString(xgr(13,i), "")'비용유형
		t_account_str = f_toString(xgr(13,i), "")'세부유형

		If t_bill_date <> "" Then
			t_bill_date = CStr(t_bill_date)
		End If

		'비용유형 지정
		arr_str = Split(t_account_str, ")")'세부유형

		For j = 0 To UBound(arr_str)
			If j = 0 Then
				slip_gubun = Replace(arr_str(j), "(", "")
			Else
				slip_account = arr_str(j)
			End If
		Next

		If slip_gubun = "비용" Then
			arr_account = Split(slip_account, "-")

			For k = 0 To UBound(arr_account)
				If k = 0 Then
					t_account = arr_account(k)
				Else
					t_account_item = arr_account(k)
				End If
			Next
		Else
			t_account = slip_account
			t_account_item = slip_account
		End If

		read_cnt = read_cnt + 1

		'값이 없는 항목은 적용 항목에서 제외 처리
		If t_org_code <> "" And t_company <> "" And t_mg_saupbu <> "" And t_account_str <> "" Then
			pay_method = "현금"
			vat_yn = "Y"
			pay_yn = "N"

			'매입세금계산서 조회
			objBuilder.Append "SELECT bill_date, owner_company, price, cost, cost_vat, "
			objBuilder.Append "	trade_name, trade_no, cost_reg_yn "
			objBuilder.Append "FROM tax_bill WHERE approve_no = '"&t_approve&"';"

			Set rsTax = DBConn.Execute(objBuilder.Tostring)
			objBuilder.Clear()

			slip_date = rsTax("bill_date")
			emp_company = rsTax("owner_company")
			price = CDbl(rsTax("price"))
			cost = CDbl(rsTax("cost"))
			cost_vat = CDbl(rsTax("cost_vat"))
			customer = rsTax("trade_name")
			customer_no = rsTax("trade_no")
			cost_reg_yn = rsTax("cost_reg_yn")'비용등록 플래그

			rsTax.Close()

			'사용조직 조회
			objBuilder.Append "SELECT org_bonbu, org_saupbu, org_team, org_reside_place, org_name "
			objBuilder.Append "FROM emp_org_mst "
			objBuilder.Append "WHERE org_code = '"&t_org_code&"';"

			Set rsOrg = DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			bonbu = rsOrg("org_bonbu")
			saupbu = f_toString(rsOrg("org_saupbu"), "")
			team = f_toString(rsOrg("org_team"), "")
			reside_place = f_toString(rsOrg("org_reside_place"), "")
			org_name = rsOrg("org_name")

			rsOrg.Close()

			'비용등록 사용자 조회
			'objBuilder.Append "SELECT emp_job, emp_name FROM emp_master WHERE emp_no = '"&t_emp_no&"';"
			objBuilder.Append "SELECT emp_no, emp_job FROM emp_master "
			objBuilder.Append "WHERE (emp_end_date IS NULL OR emp_end_date <> '' OR emp_end_date = '1900-01-01') "
			objBuilder.Append "	AND emp_org_code = '"&t_org_code&"' AND emp_name = '"&t_emp_name&"';"

			Set rsEmp = DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			emp_grade = rsEmp("emp_job")
			emp_no = rsEmp("emp_no")

			rsEmp.Close()

			If cost_reg_yn = "Y" Then'update
				objBuilder.Append "SELECT slip_date, slip_seq FROM general_cost "
				objBuilder.Append "WHERE approve_no = '"&t_approve&"';"

				Set rsGeneral = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				objBuilder.Append "UPDATE general_cost SET "
				objBuilder.Append "slip_date='"&rsGeneral("slip_date")&"', slip_seq='"&rsGeneral("slip_seq")&"', slip_gubun='"&slip_gubun&"', emp_company='"&t_owner_company&"',bonbu='"&bonbu&"',"
				objBuilder.Append "saupbu='"&saupbu&"', team='"&team&"', org_name='"&org_name&"', reside_place='"&reside_place&"', company='"&emp_company&"',"
				objBuilder.Append "account='"&t_account&"', account_item='"&t_account_item&"', pay_method='"&pay_method&"', price='"&price&"', cost='"&cost&"',"
				objBuilder.Append "vat_yn='"&vat_yn&"', cost_vat='"&cost_vat&"', customer='"&customer&"', customer_no='"&customer_no&"', emp_name='"&t_emp_name&"',"
				objBuilder.Append "emp_no='"&emp_no&"', emp_grade='"&emp_grade&"', pay_yn='"&pay_yn&"', slip_memo='"&t_tax_bill_memo&"', tax_bill_yn='Y',"
				objBuilder.Append "cancel_yn='N', end_yn='N', mod_id='"&user_id&"', mod_user='"&user_name&"', mod_date=NOW(),"
				objBuilder.Append "mg_saupbu='"&t_mg_saupbu&"' "
				objBuilder.Append "WHERE approve_no = '"&t_approve&"';"

				DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				rsGeneral.Close()
			Else'insert
				'매입계산서 목록 조회
				objBuilder.Append "SELECT MAX(slip_seq) AS 'max_seq' FROM general_cost "
				objBuilder.Append "WHERE slip_date='"&slip_date&"';"

				Set rsGe = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				If IsNull(rsGe("max_seq")) Then
					slip_seq = "001"
				Else
					max_seq = "00" & CStr((Int(rsGe("max_seq")) + 1))
					slip_seq = Right(max_seq, 3)
				End If

				rsGe.Close()

				objBuilder.Append "INSERT INTO general_cost("
				objBuilder.Append "slip_date, slip_seq, slip_gubun, emp_company,bonbu, "
				objBuilder.Append "saupbu, team, org_name, reside_place, company, "
				objBuilder.Append "account, account_item, pay_method, price, cost, "
				objBuilder.Append "vat_yn, cost_vat, customer, customer_no, emp_name, "
				objBuilder.Append "emp_no, emp_grade, pay_yn, slip_memo, tax_bill_yn, "
				objBuilder.Append "cancel_yn, end_yn, reg_id, reg_user, reg_date, "
				objBuilder.Append "approve_no, mg_saupbu "
				objBuilder.Append ")VALUES("
				objBuilder.Append "'"&slip_date&"','"&slip_seq&"','"&slip_gubun&"','"&t_owner_company&"','"&bonbu&"', "
				objBuilder.Append "'"&saupbu&"','"&team&"','"&org_name&"','"&reside_place&"','"&t_company&"', "
				objBuilder.Append "'"&t_account&"','"&t_account_item&"','"&pay_method&"',"&price&","&cost&", "
				objBuilder.Append "'"&vat_yn&"',"&cost_vat&",'"&customer&"','"&customer_no&"','"&t_emp_name&"', "
				objBuilder.Append "'"&emp_no&"','"&emp_grade&"','"&pay_yn&"','"&t_tax_bill_memo&"','Y', "
				objBuilder.Append "'N','N','"&user_id&"','"&user_name&"',NOW(), "
				objBuilder.Append "'"&t_approve&"','"&t_mg_saupbu&"');"

				DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				'sql = "Update tax_bill set cost_reg_yn='Y',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where approve_no = '"&approve_no&"'"
				objBuilder.Append "UPDATE tax_bill SET "
				objBuilder.Append "	cost_reg_yn='Y', mod_id='"&user_id&"', mod_name='"&user_name&"', mod_date = NOW() "
				objBuilder.Append "WHERE approve_no = '"&t_approve&"';"

				DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()
			End If

			'처리 개수
			write_cnt = write_cnt + 1
		End If
	Next
End If
Set rsTax = Nothing
Set rsOrg = Nothing
Set rsEmp = Nothing

If cost_reg_yn <> "Y" Then
	Set rsGe = Nothing
	Set rsGeneral = Nothing
End If

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "변경 중 Error가 발생하였습니다."
	url = "/cost/tax_esero_upload.asp"
Else
	DBConn.CommitTrans
	err_msg = "총 "&CStr(read_cnt)&"건 읽고 "&CStr(write_cnt)&" 건 처리되었습니다."

	'url = "/finance/tax_esero_mg.asp?bill_id=1&"&bill_month="&bill_month&"&cost_reg_yn=T&end_yn=T"
	url = "/cost/tax_esero_upload.asp"
End If

rs.Close() : Set rs = Nothing
cn.Close() : Set cn = Nothing
DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&err_msg&"');"
Response.Write "	location.replace('"&url&"');"
Response.Write "</script>"
Response.End
%>