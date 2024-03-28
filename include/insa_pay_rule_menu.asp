<div class="btnRight">
	<a href="/pay/insa_pay_insurance_mg.asp" class="btnType01">4대보험 요율설정</a>
<%
If SysAdminYn = "Y" Then
%>
	<a href="/pay/insa_pay_rule_mg.asp" class="btnType01"><span style="color:red;">4대보험 요율설정</span></a>
	<a href="/insa_pay_income_rule_mg.asp" class="btnType01"><span style="color:red;">근로소득세율 설정</span></a>
	<a href="/insa_pay_income_amount.asp" class="btnType01"><span style="color:red;">근로소득 간이세액</span></a>
	<a href="/insa_pay_year_rule_trans.asp" class="btnType01"><span style="color:red;">급여기초자료 이월(1월급여처리전)</span></a>
	<a href="insa_pay_bonus_mg.asp" class="btnType01"><span style="color:red;">상여율 설정</span></a>
	<a href="insa_pay_other_tax.asp" class="btnType01"><span style="color:red;">기타세액공제 설정</span></a>
<%
End If
%>
</div>
