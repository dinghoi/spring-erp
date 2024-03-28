<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'On Error Resume Next

Server.ScriptTimeOut = 1200

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
Dim org_company, end_month, end_yn, cost_year, cost_month
Dim from_date, end_date, to_date, start_date
Dim rs_oil
Dim deptName
Dim emp_msg, end_msg

org_company	=	Request("org_company")
deptName		=	Request("saupbu")	'사업 본부명으로 변경 사용
end_month	=	Request("end_month")
end_yn		=	Request("end_yn")

cost_year 	= Mid(end_month, 1, 4)
cost_month 	= Mid(end_month, 5)

from_date = Mid(end_month, 1, 4) & "-" & Mid(end_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))
start_date = DateAdd("m", -1, from_date)

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('마감 처리중!');"
Response.Write "</script>"

DBConn.BeginTrans

objBuilder.Append "SELECT oil_unit_id "
objBuilder.Append "FROM oil_unit "
objBuilder.Append "WHERE oil_unit_month = '"&end_month&"' "

Set rs_oil = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rs_oil.EOF Or rs_oil.BOF Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('유류비 단가가 입력되어 있지 않아 마감을 할 수 없습니다.');"
	Response.Write "	location.replace('/cost/cost_end_mg.asp');"
	Response.Write "</script>"
	Response.End
Else
	'유류비 단가 및 계산
%>
	<!--#include virtual="/cost/inc/inc_cost_end_oil.asp" -->
<%
	'개인 경비 정산(교통비, 야특근, 카드)
	emp_cnt = 0
%>
	<!--#include virtual="/cost/inc/inc_cost_end_person.asp" -->
<%
	'월별 인사마스터 구성 여부 파악
	If emp_cnt > 0 Then
		'4대보험 및 급여 SUM 처리
%>
		<!--#include virtual="/cost/inc/inc_cost_end_sum_insure.asp" -->
<%
		'상여/알바비 SUM 처리
%>
		<!--#include virtual="/cost/inc/inc_cost_end_sum_bonus.asp" -->
<%
		'DB SUM 일반 경비
%>
		<!--#include virtual="/cost/inc/inc_cost_end_sum_cost.asp" -->
<%
		'DB SUM 교통비
%>
		<!--#include virtual="/cost/inc/inc_cost_end_sum_transit.asp" -->
<%
		'카드비용 집계
%>
		<!--#include virtual="/cost/inc/inc_cost_end_sum_card.asp" -->
<%
		'cost_end 테이블의 saupbu 컬럼을 본부명과 매칭 사용[허정호_20210312]
		If end_yn = "C" Then
			objBuilder.Append "UPDATE cost_end SET "
			objBuilder.Append "	end_yn = 'Y', mod_id = '"&user_id&"', mod_name = '"&user_name&"', mod_date = NOW() "
			objBuilder.Append "WHERE end_month = '"&end_month&"' "
			objBuilder.Append "	AND saupbu = '"&deptName&"' "
		Else
			objBuilder.Append "DELETE FROM cost_end "
			objBuilder.Append "WHERE end_month = '"&end_month&"' AND saupbu = '"&deptName&"' "

			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			objBuilder.Append "INSERT INTO cost_end(end_month, saupbu, end_yn, batch_yn, bonbu_yn, ceo_yn, reg_id, reg_name, reg_date)"
			objBuilder.Append "VALUES("
			objBuilder.Append "'"&end_month&"', '"&deptName&"', 'Y', 'N', 'N', 'N', '"&user_id&"', '"&user_name&"', NOW()) "
		End If

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	End If
	' 월별 인사마스터 구성 여부 파악 END

	If emp_cnt = 0 Then
		emp_msg = "인사마스터 마감이 되지 않았습니다."
	Else
		emp_msg = ""
	End If

	If Err.Number <> 0 Then
		DBConn.RollbackTrans
		end_msg = emp_msg & "처리중 Error가 발생하였습니다."
	Else
		DBConn.CommitTrans
		end_msg = emp_msg & "마감처리 되었습니다."
	End If

	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('"&end_msg&"');"
	Response.Write "	location.replace('/cost/cost_end_mg.asp');"
	Response.Write "</script>"
	Response.End
End If
rs_oil.Close() : Set rs_oil = Nothing
DBConn.Close() : Set DBConn = Nothing
%>