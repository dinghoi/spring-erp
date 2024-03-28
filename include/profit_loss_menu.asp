<div class="btnRight">
	<a href="/sales/reside_cost_report.asp" class="btnType01">비용유형별 현황</a>
	<a href="/sales/saupbu_profit_loss_total.asp" class="btnType01">사업부별 손익총괄</a>

<%If (org_name = "회계재무" And account_grade = "0") Or SysAdminYn = "Y" Or empProfitGrade = "Y" Then%>
	<a href="/sales/management_cost_report.asp" class="btnType01">전사공통비</a>
<%End If%>

<%If (org_name = "회계재무" And account_grade = "0") Or SysAdminYn = "Y" Or (empProfitGrade = "Y" And partCostView = "Y") Then%>
	<a href="/sales/part_cost_report.asp" class="btnType01">부문공통비</a>
<%End If%>

<%If empProfitViewAll = "Y" Or SysAdminYn = "Y" Or (empProfitGrade = "Y" And subPartCostView = "Y") Then%>
	<a href="/sales/saupbu_ksys_part_cost.asp" class="btnType01">부문공통비(2)</a>
<%End If%>

<%'If CoworkYn = "Y" Or SysAdminYn = "Y" Or empProfitGrade = "Y" Then%>
<%If CompanyCostYn = "Y" Or SysAdminYn = "Y"  Then %>
	<a href="/sales/company_cost_report.asp" class="btnType01">거래처별 손익현황</a>
<%End If %>

<%'If CoworkYn = "Y" Or SysAdminYn = "Y" Or empProfitGrade = "Y" Then%>
<%If CoworkYn = "Y" Or SysAdminYn = "Y" Then%>
	<a href="/sales/company_cost_cowork.asp" class="btnType01">거래처별 협업</a>
<%End If%>

<%If cost_grade = "0" Then%>
	<a href="/saupbu_emp_report.asp" class="btnType01">사업부별 인원현황</a>
<%End If%>

<%If empReportKDCYn = "Y" Or SysAdminYn = "Y" Then%>
	<a href="/sales/saupbu_emp_report_kdc.asp" class="btnType01">사업부별 인원현황(KDC)</a>
<%End If%>

<%If SysAdminYn = "Y" Then%>
	<a href="/sales/part_company_report.asp" class="btnType01">장애처리 현황(미노출)</a>
	<a href="/part_cost_report2.asp" class="btnType01">AS 배부기준(변경후, 미노출)</a>
	<a href="management_cost_report2.asp" class="btnType01">전사공통비배부기준(변경후, 미노출)</a>
	<a href="/saupbu_profit_loss_month_20200408.asp" class="btnType01">사업부별 월별손익(수정, 미노출)</a>
	<a href="/saupbu_profit_loss_month.asp" class="btnType01">사업부별 월별손익(미노출)</a>
	<a href="/company_profit_loss_report.asp" class="btnType01">고객사별 손익현황(미노출)</a>
	<a href="/sales/saupbu_profit_loss_total_std.asp" class="btnType01">사업부별 손익총괄(표준, 미노출)</a>
	<a href="/sales/part_cost_report_unit.asp" class="btnType01">부문공통비 AS 배부기준(미노출)</a>
	<a href="/sales/old/saupbu_profit_loss_total_old.asp" class="btnType01">사업부별 손익총괄(구버전, 미노출)</a>
<%End If%>

</div>
