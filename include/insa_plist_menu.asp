<div class="btnRight">
	<a href="/person/insa_plist_pay_mg.asp" class="btnType01">급여 현황</a>

<%' if (position = "총괄대표") or (position = "본부장") or (position = "사업부장") Or (in_empno = "900002") then %>
<%If sales_grade < "2" Then %>
	<a href="/person/insa_manager_emp_list.asp" class="btnType01">직원 현황</a>
<%End If %>

<%If SysAdminYn = "Y" Then%>
	<a href="/person/insa_plist_mg.asp" class="btnType01"><span style="color:red;">직원 주소록</span></a>
<%End If%>
</div>
