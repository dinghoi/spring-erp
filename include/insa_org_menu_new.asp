<div class="btnRight">
<%
Dim in_name, in_empno

in_name = request.cookies("nkpmg_user")("coo_user_name")
in_empno = request.cookies("nkpmg_user")("coo_user_id")
position = request.cookies("nkpmg_user")("coo_position")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
%>
	<a href="./insa_org_mg.asp" class="btnType01">조직현황</a>
	<a href="../insa_org_name_view.asp" class="btnType01">조직명 조회</a>
	<a href="../insa_org_end.asp" class="btnType01">조직폐쇄</a>
	<a href="../insa_org_to_list.asp" class="btnType01">조직별 T.O현황</a>
	<a href="../insa_org_list.asp" class="btnType01">조직 조건조회</a>
	<a href="../insa_emp_juso_list.asp" class="btnType01">직원주소록</a>
	<a href="../insa_stay_mg.asp" class="btnType01">실근무지 관리</a>
	<a href="../insa_etc_code_mg.asp" class="btnType01">인사 코드관리</a>

	<a href="#" onClick="pop_Window('../insa_user_password.asp?u_type=<%="U"%>','insa_user_password_pop','scrollbars=yes,width=500,height=350')" class="btnType01">사용자비밀번호 확인</a>

	<a href="../insa_mg_list.asp" class="btnType01">인사자료미등록현황</a>
<%'//2017-08-14 권한자 수정 : 이윤영대리(100104), 윤성희사원(101063) 추가, 박진성(101168) 삭제 %>
<%If in_empno = "100104" Or in_empno = "101063" Or in_empno="101100" Or in_empno="100018" Or in_empno="101622" Or in_empno="102560" Or in_empno = "102592" Then %>
<%'<a href="insa_org_mst_month_save.asp" class="btnType01">월 마감</a> %>
	<a href="#" onClick="pop_Window('../insa_month_final_submit.asp','insa_month_final_pop','scrollbars=yes,width=750,height=350')" class="btnType01">월 마감</a>
<%End If%>
</div>
