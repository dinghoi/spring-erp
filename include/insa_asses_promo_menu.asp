<%
Dim in_name, in_empno

in_name = Request.Cookies("nkpmg_user")("coo_user_name")
in_empno = Request.Cookies("nkpmg_user")("coo_user_id")
'position = request.cookies("nkpmg_user")("coo_position")
%>
<div class="btnRight">
	<a href="/insa/insa_promotion_list.asp" class="btnType01">�����������Ȳ</a>

<!--�̻�� �޴� �ӽ� �ּ� ó��[����ȣ_20210402]-->
<%'If in_empno = "102592" Then %>
	<!--<a href="/insa_emp_owner_org_list.asp" class="btnType01">������������</a>-->
<%'End If %>

<!--�̻�� �޴� �ӽ� �ּ� ó��[����ȣ_20210402]-->
<%'If in_empno = "102592" Then %>
	<!--<a href="/insa_pay_total_info.asp" class="btnType01" target="_parent">����κ� �ΰǺ���ȸ</a>-->
<%'End If %>
	<a href="/insa/insa_emp_master_mg.asp" class="btnType01" target="_parent">������ ����</a>
</div>
