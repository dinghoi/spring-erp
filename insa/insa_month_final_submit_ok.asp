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
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim be_month, pre_month, rsPayCnt, payCnt
Dim be_yyyy, be_mm, end_msg

'emp_user = request.cookies("nkpmg_user")("coo_user_name")
'be_month = request.form("be_month")

be_month = f_Request("inc_yyyy1")	'마감 월
pre_month = f_Request("pre_month")	'이전 월

be_yyyy = CStr(Mid(be_month, 1, 4))
be_mm = CStr(Mid(be_month, 5, 6))

'급여 마감 여부
objBuilder.Append "SELECT COUNT(*) FROM pay_month_give "
objBuilder.Append "WHERE pmg_yymm = '"&be_month&"' "

Set rsPayCnt = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

payCnt = rsPaycnt(0)

rsPayCnt.Close() : Set rsPayCnt = Nothing

If CInt(payCnt) = 0 Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('"&be_yyyy&"년 "&be_mm&"월 급여 마감 후 진행 가능합니다.');"
	Response.Write "	history.go(-1);"
	Response.Write "</script>"
	Response.End
End If

DBConn.BeginTrans

'월 조직 정보 처리
objBuilder.Append "DELETE FROM emp_org_mst_month WHERE org_month ='"&be_month&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

objBuilder.Append "INSERT INTO emp_org_mst_month  "
objBuilder.Append "SELECT '"&be_month&"' AS org_month, emp_org_mst.* "
objBuilder.Append "FROM emp_org_mst "
objBuilder.Append "WHERE emp_org_mst.org_end_date IS NULL OR emp_org_mst.org_end_date = '0000-00-00' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

Dim rsPreEmp, arrPreEmp, i

'전월 인사 정보 조회
objBuilder.Append "SELECT emmt.emp_no, emmt.cost_group, emmt.cost_center, "
objBuilder.Append "	emmt.emp_company, emmt.emp_bonbu, emmt.emp_saupbu, emmt.emp_team, "
objBuilder.Append "	emmt.emp_org_code, emmt.emp_org_name, "
objBuilder.Append "	pmgt.cost_group, pmgt.cost_center, "
objBuilder.Append "	pmgt.pmg_company, pmgt.pmg_bonbu, pmgt.pmg_team, "
objBuilder.Append "	pmgt.pmg_org_name, pmgt.mg_saupbu, pmgt.pmg_id, pmgt.pmg_saupbu "
objBuilder.Append "FROM pay_month_give AS pmgt "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
objBuilder.Append "	AND emmt.emp_month = '"&pre_month&"' "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emmt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE pmgt.pmg_id = '1' "
objBuilder.Append "	AND pmgt.pmg_yymm = '"&pre_month&"'"

Set rsPreEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsPreEmp.EOF Then
	arrPreEmp = rsPreEmp.getRows()
End If

rsPreEmp.Close() : Set rsPreEmp = Nothing

Dim cost_group, cost_center, emp_bonbu, emp_saupbu
Dim emp_team, emp_org_code, emp_org_name, pmg_cost_group, pmg_cost_center
Dim pmg_company, pmg_bonbu, pmg_team, pmg_org_name, mg_saupbu, pmg_id, pmg_saupbu

'월 직원/급여 정보 업데이트(전월 정보 업데이트 처리)
If IsArray(arrPreEmp) Then
	For i = LBound(arrPreEmp) To UBound(arrPreEmp, 2)
		emp_no = arrPreEmp(0, i)
		cost_group = arrPreEmp(1, i)
		cost_center = arrPreEmp(2, i)
		emp_company = arrPreEmp(3, i)
		emp_bonbu = arrPreEmp(4, i)
		emp_saupbu = arrPreEmp(5, i)
		emp_team = arrPreEmp(6, i)
		emp_org_code = arrPreEmp(7, i)
		emp_org_name = arrPreEmp(8, i)
		pmg_cost_group = arrPreEmp(9, i)
		pmg_cost_center = arrPreEmp(10, i)
		pmg_company = arrPreEmp(11, i)
		pmg_bonbu = arrPreEmp(12, i)
		pmg_team = arrPreEmp(13, i)
		pmg_org_name = arrPreEmp(14, i)
		mg_saupbu = arrPreEmp(15, i)
		pmg_id = arrPreEmp(16, i)
		pmg_saupbu = arrPreEmp(17, i)

		objBuilder.Append "UPDATE emp_master_month SET "
		objBuilder.Append "	cost_group = '"&cost_group&"', cost_center = '"&cost_center&"', emp_company = '"&emp_company&"', "
		objBuilder.Append "	emp_bonbu = '"&emp_bonbu&"', emp_saupbu = '"&emp_saupbu&"', emp_team = '"&emp_team&"', "
		objBuilder.Append "	emp_org_code = '"&emp_org_code&"', emp_org_name = '"&emp_org_name&"', "
		objBuilder.Append "	emp_mod_date = NOW(), emp_mod_user = '"&user_id&"' "
		objBuilder.Append "WHERE emp_month = '"&be_month&"' AND emp_no = '"&emp_no&"';"

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		objBuilder.Append "UPDATE pay_month_give SET "
		objBuilder.Append "	cost_group = '"&pmg_cost_group&"', cost_center = '"&pmg_cost_center&"', pmg_company = '"&pmg_company&"', "
		objBuilder.Append "	pmg_bonbu = '"&pmg_bonbu&"', pmg_team = '"&pmg_team&"', pmg_org_name = '"&pmg_org_name&"', "
		objBuilder.Append "	mg_saupbu = '"&mg_saupbu&"', pmg_mod_date = NOW(), pmg_mod_user = '"&user_id&"' "
		objBuilder.Append "WHERE pmg_yymm = '"&be_month&"' AND pmg_emp_no = '"&emp_no&"' "
		objBuilder.Append "	AND pmg_id = '"&pmg_id&"' AND pmg_company = '"&pmg_company&"';"

		'Response.write objBuilder.Tostring() & "<br/>"

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "처리중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = be_month&" 조직 및 인사 마스타 마감처리가 되었습니다."

	'Response.Write "<script type='text/javascript'>"
	'Response.Write "	alert('"&be_month&" 조직 및 인사 마스타 마감처리가 되었습니다.');"
	'Response.Write "	window.close();"
	'Response.Write "</script>"
	'Response.End
End If

Response.Write end_msg

DBConn.Close() : Set DBConn = Nothing
Response.End
%>
