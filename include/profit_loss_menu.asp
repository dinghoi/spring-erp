<div class="btnRight">
	<a href="/sales/reside_cost_report.asp" class="btnType01">��������� ��Ȳ</a>
	<a href="/sales/saupbu_profit_loss_total.asp" class="btnType01">����κ� �����Ѱ�</a>

<%If (org_name = "ȸ���繫" And account_grade = "0") Or SysAdminYn = "Y" Or empProfitGrade = "Y" Then%>
	<a href="/sales/management_cost_report.asp" class="btnType01">��������</a>
<%End If%>

<%If (org_name = "ȸ���繫" And account_grade = "0") Or SysAdminYn = "Y" Or (empProfitGrade = "Y" And partCostView = "Y") Then%>
	<a href="/sales/part_cost_report.asp" class="btnType01">�ι������</a>
<%End If%>

<%If empProfitViewAll = "Y" Or SysAdminYn = "Y" Or (empProfitGrade = "Y" And subPartCostView = "Y") Then%>
	<a href="/sales/saupbu_ksys_part_cost.asp" class="btnType01">�ι������(2)</a>
<%End If%>

<%'If CoworkYn = "Y" Or SysAdminYn = "Y" Or empProfitGrade = "Y" Then%>
<%If CompanyCostYn = "Y" Or SysAdminYn = "Y"  Then %>
	<a href="/sales/company_cost_report.asp" class="btnType01">�ŷ�ó�� ������Ȳ</a>
<%End If %>

<%'If CoworkYn = "Y" Or SysAdminYn = "Y" Or empProfitGrade = "Y" Then%>
<%If CoworkYn = "Y" Or SysAdminYn = "Y" Then%>
	<a href="/sales/company_cost_cowork.asp" class="btnType01">�ŷ�ó�� ����</a>
<%End If%>

<%If cost_grade = "0" Then%>
	<a href="/saupbu_emp_report.asp" class="btnType01">����κ� �ο���Ȳ</a>
<%End If%>

<%If empReportKDCYn = "Y" Or SysAdminYn = "Y" Then%>
	<a href="/sales/saupbu_emp_report_kdc.asp" class="btnType01">����κ� �ο���Ȳ(KDC)</a>
<%End If%>

<%If SysAdminYn = "Y" Then%>
	<a href="/sales/part_company_report.asp" class="btnType01">���ó�� ��Ȳ(�̳���)</a>
	<a href="/part_cost_report2.asp" class="btnType01">AS ��α���(������, �̳���)</a>
	<a href="management_cost_report2.asp" class="btnType01">���������α���(������, �̳���)</a>
	<a href="/saupbu_profit_loss_month_20200408.asp" class="btnType01">����κ� ��������(����, �̳���)</a>
	<a href="/saupbu_profit_loss_month.asp" class="btnType01">����κ� ��������(�̳���)</a>
	<a href="/company_profit_loss_report.asp" class="btnType01">���纰 ������Ȳ(�̳���)</a>
	<a href="/sales/saupbu_profit_loss_total_std.asp" class="btnType01">����κ� �����Ѱ�(ǥ��, �̳���)</a>
	<a href="/sales/part_cost_report_unit.asp" class="btnType01">�ι������ AS ��α���(�̳���)</a>
	<a href="/sales/old/saupbu_profit_loss_total_old.asp" class="btnType01">����κ� �����Ѱ�(������, �̳���)</a>
<%End If%>

</div>
