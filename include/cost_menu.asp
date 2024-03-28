<div class="btnRight">
	<a href="/person_cost_report.asp" class="btnType01">개인비용정산</a>
	<a href="/general_cost_mg.asp" class="btnType01">일반경비</a>
<%If cost_grade="0" Or cost_grade="3" Then %>
	<a href="/others_cost_mg.asp" class="btnType01">비용대행</a>
<%End If%>
<%If cost_grade="6" Or cost_grade="5" Or cost_grade<"3" Then%>
	<a href="/overtime_mg.asp" class="btnType01">야특근</a>
<%End If%>
<%If cost_grade="6" Or cost_grade="5" Or cost_grade<"3" Then%>
	<a href="/transit_cost_mg.asp" class="btnType01">교통비</a>
	<!--<a href="/cost/transit_cost_mg.asp" class="btnType01">교통비</a>-->
<%End If%>
	<a href="/person_card_mg.asp" class="btnType01">개인별카드내역</a>
<%If cost_grade<"6" Then %>
	<a href="/alba_cost_mg.asp" class="btnType01">아르바이트</a>
	<!--<a href="/tax_esero_in_mg.asp" class="btnType01">E세로매입세금계산서</a>-->
	<a href="/cost/tax_esero_in_mg.asp" class="btnType01">E세로매입세금계산서</a>
<%If account_grade = "0" Then%>
	<a href="/cost/tax_esero_upload.asp" class="btnType01">E세로비용일괄업로드</a>
<%End If%>
	<a href="/cost/tax_bill_in_mg.asp" class="btnType01">매입세금계산서</a>
	<a href="/tax_bill_manual_mg.asp" class="btnType01">종이세금계산서</a>
<%End If%>
<%If cost_grade="0" Then %>
	<a href="/cost/depreciation_cost_mg.asp" class="btnType01">상각비</a>
<%End If%>
</div>
