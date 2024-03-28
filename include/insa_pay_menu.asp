<div class="btnRight">
	<a href="/pay/insa_pay_mg.asp" class="btnType01">급여지급현황</a>
	<a href="/pay/insa_pay_month_batch.asp" class="btnType01">급여기초이월</a>
	<a href="/pay/insa_pay_month_pay_mg.asp" class="btnType01">급여입력</a>
	<a href="/pay/insa_pay_month_up.asp" class="btnType01">급여Upload</a>
	<a href="/pay/insa_pay_bank_transfer.asp" class="btnType01">은행이체명세</a>
	<a href="/pay/insa_pay_month_ledger.asp" class="btnType01">월 급여대장</a>
<%
If SysAdminYn = "Y" Then
%>
	<a href="insa_pay_overtime_report.asp" class="btnType01"><span style="color:red;">야·특근수당</span></a>
	<a href="/insa_pay_overtime_report2.asp" class="btnType01"><span style="color:red;">야·특근수당</span></a>
	<a href="insa_pay_overtime_report3.asp" class="btnType01"><span style="color:red;">야·특근수당</span></a>
	<a href="/insa_pay_sawo_report.asp" class="btnType01"><span style="color:red;">경조금공제</span></a>
	<a href="insa_system_popup.asp" class="btnType01"><span style="color:red;">급여자료일괄입력</span></a>
	<a href="insa_pay_expense_mg.asp" class="btnType01"><span style="color:red;">지급/공제입력</span></a>
	<a href="/insa_pay_msum_report.asp" class="btnType01"><span style="color:red;">월급여 항목별집계</span></a>
	<a href="/insa_pay_incentive_mg.asp" class="btnType01"><span style="color:red;">상여금등</span></a>
	<a href="/insa_pay_incentive_up.asp" class="btnType01"><span style="color:red;">상여금Upload</span></a>
	<a href="/insa_pay_month_saup_list.asp" class="btnType01"><span style="color:red;">조직별 급여내역</span></a>
	<a href="/insa_pay_comment_list.asp" class="btnType01"><span style="color:red;">급여특이사항</span></a>
<%
End If
%>
</div>
