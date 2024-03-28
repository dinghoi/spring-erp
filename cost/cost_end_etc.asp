<%
' 유류비 단가 및 유류비 계산
%>
<!--#include virtual="/cost/inc/inc_bonbu_end_oil.asp" -->
<%
' 개인별 비용 정산
%>
<!--#include virtual="/cost/inc/inc_bonbu_end_person.asp" -->
<%
' 월별 인사마스터 구성 여부 파악
If emp_cnt > 0 Then
	'4대보험 및 급여 SUM 처리
%>
	<!--#include virtual="/cost/inc/inc_bonbu_end_sum_insure.asp" -->
<%
	'상여/알바비 SUM 처리
%>
	<!--#include virtual="/cost/inc/inc_bonbu_end_sum_bonus.asp" -->
<%
	'DB SUM 일반 경비
%>
	<!--#include virtual="/cost/inc/inc_bonbu_end_sum_cost.asp" -->
<%
	'DB SUM 교통비
%>
	<!--#include virtual="/cost/inc/inc_bonbu_end_sum_transit.asp" -->
<%
	'카드비용 집계
%>
	<!--#include virtual="/cost/inc/inc_bonbu_end_sum_card.asp" -->
<%
	objBuilder.Append "CALL USP_ORG_END_PROC('"&end_month&"', '사업부외나머지', '"&end_yn&"', '"&user_id&"', '"&user_name&"');"
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If
' 월별 인사마스터 구성 여부 파악 END