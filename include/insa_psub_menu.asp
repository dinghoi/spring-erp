<div class="btnRight">
<%If IntroMemberYn = "N" Then%>
	<a href="#" onClick="pop_Window('/person/insa_individual_card00.asp?emp_no=<%=in_empno%>','emp_card0_pop','scrollbars=yes,width=1300,height=650')" class="btnType01">인사기록카드</a>
	<a href="/person/insa_individual_emp_add.asp" class="btnType01">인사기본수정</a>
	<a href="/person/insa_individual_family.asp" class="btnType01">가족사항</a>
	<a href="/person/insa_individual_school.asp" class="btnType01">학력사항</a>
	<a href="/person/insa_individual_career.asp" class="btnType01">경력사항</a>
	<a href="/person/insa_individual_qual.asp" class="btnType01">자격사항</a>
	<a href="/person/insa_individual_edu.asp" class="btnType01">교육사항</a>
	<a href="/person/insa_individual_language.asp" class="btnType01">어학능력</a>
<%Else%>
	<a href="/member/member_add.asp" class="btnType01">회원가입</a>
	<a href="/member/member_family.asp" class="btnType01">가족사항</a>
	<a href="/member/member_school.asp" class="btnType01">학력사항</a>
	<a href="/member/member_career.asp" class="btnType01">경력사항</a>
	<a href="/member/member_qual.asp" class="btnType01">자격사항</a>
	<a href="/member/member_edu.asp" class="btnType01">교육사항</a>
	<a href="/member/member_language.asp" class="btnType01">어학능력</a>
<%End If%>
</div>
