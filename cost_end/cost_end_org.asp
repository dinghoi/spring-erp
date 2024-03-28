<%
'전체 본부 별 비용 마감 처리 추가(일괄 정산용으로 추가)[허정호_20211007]
objBuilder.Append "CALL USP_ORG_END_SALES_SEL();"
Set rsSalesOrg = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsSalesOrg.EOF Then
	arrSalesOrg = rsSalesOrg.getRows()
End If
rsSalesOrg.Close() : Set rsSalesOrg = Nothing

If IsArray(arrSalesOrg) Then
	For oLoop = LBound(arrSalesOrg) To UBound(arrSalesOrg, 2)
		deptName = arrSalesOrg(0, oLoop)	'본부명
		emp_cnt = 0

		'유류비 단가 및 계산
%>
		<!--#include virtual="/cost/inc/inc_org_end_oil.asp" -->
<%
		'개인 경비 정산(교통비, 야특근, 카드)
%>
		<!--#include virtual="/cost/inc/inc_org_end_person.asp" -->
<%
		'월별 인사마스터 구성 여부 파악
		If emp_cnt > 0 Then
			'4대보험 및 급여 SUM 처리
%>
			<!--#include virtual="/cost/inc/inc_org_end_sum_insure.asp" -->
<%
			'상여/알바비 SUM 처리
%>
			<!--#include virtual="/cost/inc/inc_org_end_sum_bonus.asp" -->
<%
			'DB SUM 일반 경비
%>
			<!--#include virtual="/cost/inc/inc_org_end_sum_general.asp" -->
<%
			'DB SUM 교통비
%>
			<!--#include virtual="/cost/inc/inc_org_end_sum_transit.asp" -->
<%
			'카드비용 집계
%>
			<!--#include virtual="/cost/inc/inc_org_end_sum_card.asp" -->
<%
			'cost_end 테이블의 saupbu 컬럼을 본부명과 매칭 사용[허정호_20210312]
			objBuilder.Append "CALL USP_ORG_END_PROC('"&end_month&"', '"&deptName&"', '"&end_yn&"', '"&user_id&"', '"&user_name&"');"
			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
		End If
		' 월별 인사마스터 구성 여부 파악 END
	Next
End If
%>