<div class="btnRight">
	<a href="/person/insa_plist_pay_mg.asp" class="btnType01">�޿� ��Ȳ</a>

<%' if (position = "�Ѱ���ǥ") or (position = "������") or (position = "�������") Or (in_empno = "900002") then %>
<%If sales_grade < "2" Then %>
	<a href="/person/insa_manager_emp_list.asp" class="btnType01">���� ��Ȳ</a>
<%End If %>

<%If SysAdminYn = "Y" Then%>
	<a href="/person/insa_plist_mg.asp" class="btnType01"><span style="color:red;">���� �ּҷ�</span></a>
<%End If%>
</div>
