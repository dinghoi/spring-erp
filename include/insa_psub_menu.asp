<div class="btnRight">
<%If IntroMemberYn = "N" Then%>
	<a href="#" onClick="pop_Window('/person/insa_individual_card00.asp?emp_no=<%=in_empno%>','emp_card0_pop','scrollbars=yes,width=1300,height=650')" class="btnType01">�λ���ī��</a>
	<a href="/person/insa_individual_emp_add.asp" class="btnType01">�λ�⺻����</a>
	<a href="/person/insa_individual_family.asp" class="btnType01">��������</a>
	<a href="/person/insa_individual_school.asp" class="btnType01">�з»���</a>
	<a href="/person/insa_individual_career.asp" class="btnType01">��»���</a>
	<a href="/person/insa_individual_qual.asp" class="btnType01">�ڰݻ���</a>
	<a href="/person/insa_individual_edu.asp" class="btnType01">��������</a>
	<a href="/person/insa_individual_language.asp" class="btnType01">���дɷ�</a>
<%Else%>
	<a href="/member/member_add.asp" class="btnType01">ȸ������</a>
	<a href="/member/member_family.asp" class="btnType01">��������</a>
	<a href="/member/member_school.asp" class="btnType01">�з»���</a>
	<a href="/member/member_career.asp" class="btnType01">��»���</a>
	<a href="/member/member_qual.asp" class="btnType01">�ڰݻ���</a>
	<a href="/member/member_edu.asp" class="btnType01">��������</a>
	<a href="/member/member_language.asp" class="btnType01">���дɷ�</a>
<%End If%>
</div>
