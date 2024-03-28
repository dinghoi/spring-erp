<%
Dim in_name, in_empno

in_name = Request.Cookies("nkpmg_user")("coo_user_name")
in_empno = Request.Cookies("nkpmg_user")("coo_user_id")
'position = request.cookies("nkpmg_user")("coo_position")
%>
<div class="btnRight">
	<a href="/insa/insa_promotion_list.asp" class="btnType01">승진대상자현황</a>

<!--미사용 메뉴 임시 주석 처리[허정호_20210402]-->
<%'If in_empno = "102592" Then %>
	<!--<a href="/insa_emp_owner_org_list.asp" class="btnType01">상위조직변경</a>-->
<%'End If %>

<!--미사용 메뉴 임시 주석 처리[허정호_20210402]-->
<%'If in_empno = "102592" Then %>
	<!--<a href="/insa_pay_total_info.asp" class="btnType01" target="_parent">사업부별 인건비조회</a>-->
<%'End If %>
	<a href="/insa/insa_emp_master_mg.asp" class="btnType01" target="_parent">직원별 관리</a>
</div>
