<div class="btnRight">
<%
'in_name = request.cookies("nkpmg_user")("coo_user_name")
'in_empno = request.cookies("nkpmg_user")("coo_user_id")
'position = request.cookies("nkpmg_user")("coo_position")
'insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
%>
    <a href="/insa/insa_report_mg.asp" class="btnType01" target="_parent">������ȸ</a>
    <a href="/insa/insa_qual_list.asp" class="btnType01" target="_parent">�ڰ���</a>
    <a href="/insa/insa_career_list.asp" class="btnType01" target="_parent">���</a>
    <a href="/insa/insa_school_list.asp" class="btnType01" target="_parent">�з�</a>
    <a href="/insa/insa_report_emp_in.asp" class="btnType01" target="_parent">�Ի��� ��Ȳ</a>
    <a href="/insa/insa_mg_list.asp" class="btnType01" target="_parent">�λ��ڷ�̵����Ȳ</a>
    <a href="/insa/insa_report_emp_out.asp" class="btnType01" target="_parent">����� ��Ȳ</a>
    <a href="/insa/insa_emp_end_list.asp" class="btnType01" target="_parent">��������ȸ</a>
    <a href="/insa/insa_emp_infor_main.asp" class="btnType01" target="_parent">�λ�������ȸ</a>
<%If SysAdminYn = "Y" Then%>
    <a href="insa_grade_count.asp" class="btnType01" target="_parent"><span style='color:red;'>���޺� ��Ȳ</span></a>
    <a href="insa_disabled_list.asp" class="btnType01" target="_parent"><span style='color:red;'>����� ��Ȳ</span></a>
    <a href="insa_report_change.asp" class="btnType01" target="_parent"><span style='color:red;'>�ο����� ��Ȳ</span></a>
    <a href="insa_report_gunsok.asp" class="btnType01" target="_parent"><span style='color:red;'>�ټ� ��Ȳ</span></a>
    <a href="insa_age_count_org.asp" class="btnType01" target="_parent"><span style='color:red;'>���ɺ� ����</span></a>
    <a href="insa_academic_count_org.asp" class="btnType01" target="_parent"><span style='color:red;'>�зº� ����</span></a>
    <a href="insa_area_count_org.asp" class="btnType01" target="_parent"><span style='color:red;'>������ ����</span></a>
    <a href="insa_month_count_mg.asp" class="btnType01" target="_parent"><span style='color:red;'>�����ο�</span></a>
    <a href="insa_report_grade.asp" class="btnType01" target="_parent"><span style='color:red;'>������ ��Ȳ</span></a>
<%End If%>
</div>