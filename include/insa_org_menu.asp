<div class="btnRight">
<%
Dim in_name, in_empno

in_name = Request.Cookies("nkpmg_user")("coo_user_name")
in_empno = Request.Cookies("nkpmg_user")("coo_user_id")
'position = request.cookies("nkpmg_user")("coo_position")
'insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
%>
    <a href="/insa/insa_org_mg.asp" class="btnType01">������Ȳ</a>
    <a href="/insa/insa_org_name_view.asp" class="btnType01">������ ��ȸ</a>
    <a href="/insa/insa_etc_code_mg.asp" class="btnType01">�λ� �ڵ����</a>

<%'//2017-08-14 ������ ���� : �������븮(100104), ��������(101063) ����, ������(101168) ����, �����(101100) ����, %>
<%'If in_empno = "100104" Or in_empno="100018" Or in_empno="101622" Or in_empno="102560" Or in_empno = "102592" Then %>

    <a href="#" onClick="pop_Window('/member/insa_user_password.asp?u_type=U','insa_user_password_pop','scrollbars=yes,width=500,height=350')" class="btnType01">����ں�й�ȣ Ȯ��</a>
<%
If InsaMasterModYn = "Y" Then
%>
    <!--<a href="insa_org_mst_month_save.asp" class="btnType01">�� ����</a>-->
    <a href="#" onClick="pop_Window('/insa/insa_month_final_submit.asp','insa_month_final_pop','scrollbars=yes,width=750,height=350')" class="btnType01">�� ����</a>
<%
End If

If SysAdminYn = "Y" Then
%>
    <a href="/insa_org_end.asp" class="btnType01"><span style="color:red;">�������</span></a>
    <a href="/insa_org_to_list.asp" class="btnType01"><span style="color:red;">������ T.O��Ȳ</span></a>
    <a href="/insa_org_list.asp" class="btnType01"><span style="color:red;">���� ������ȸ</span></a>
    <a href="/insa_emp_juso_list.asp" class="btnType01"><span style="color:red;">�����ּҷ�</span></a>
    <a href="/insa_stay_mg.asp" class="btnType01"><span style="color:red;">�Ǳٹ��� ����</span></a>
<%
End If
%>
</div>
