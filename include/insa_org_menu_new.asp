<div class="btnRight">
<%
Dim in_name, in_empno

in_name = request.cookies("nkpmg_user")("coo_user_name")
in_empno = request.cookies("nkpmg_user")("coo_user_id")
position = request.cookies("nkpmg_user")("coo_position")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
%>
	<a href="./insa_org_mg.asp" class="btnType01">������Ȳ</a>
	<a href="../insa_org_name_view.asp" class="btnType01">������ ��ȸ</a>
	<a href="../insa_org_end.asp" class="btnType01">�������</a>
	<a href="../insa_org_to_list.asp" class="btnType01">������ T.O��Ȳ</a>
	<a href="../insa_org_list.asp" class="btnType01">���� ������ȸ</a>
	<a href="../insa_emp_juso_list.asp" class="btnType01">�����ּҷ�</a>
	<a href="../insa_stay_mg.asp" class="btnType01">�Ǳٹ��� ����</a>
	<a href="../insa_etc_code_mg.asp" class="btnType01">�λ� �ڵ����</a>

	<a href="#" onClick="pop_Window('../insa_user_password.asp?u_type=<%="U"%>','insa_user_password_pop','scrollbars=yes,width=500,height=350')" class="btnType01">����ں�й�ȣ Ȯ��</a>

	<a href="../insa_mg_list.asp" class="btnType01">�λ��ڷ�̵����Ȳ</a>
<%'//2017-08-14 ������ ���� : �������븮(100104), ��������(101063) �߰�, ������(101168) ���� %>
<%If in_empno = "100104" Or in_empno = "101063" Or in_empno="101100" Or in_empno="100018" Or in_empno="101622" Or in_empno="102560" Or in_empno = "102592" Then %>
<%'<a href="insa_org_mst_month_save.asp" class="btnType01">�� ����</a> %>
	<a href="#" onClick="pop_Window('../insa_month_final_submit.asp','insa_month_final_pop','scrollbars=yes,width=750,height=350')" class="btnType01">�� ����</a>
<%End If%>
</div>
