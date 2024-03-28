<div class="btnRight">
<%
'in_name = request.cookies("nkpmg_user")("coo_user_name")
'in_empno = request.cookies("nkpmg_user")("coo_user_id")
'position = request.cookies("nkpmg_user")("coo_position")
'insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
%>
    <a href="/insa/insa_report_mg.asp" class="btnType01" target="_parent">직원조회</a>
    <a href="/insa/insa_qual_list.asp" class="btnType01" target="_parent">자격증</a>
    <a href="/insa/insa_career_list.asp" class="btnType01" target="_parent">경력</a>
    <a href="/insa/insa_school_list.asp" class="btnType01" target="_parent">학력</a>
    <a href="/insa/insa_report_emp_in.asp" class="btnType01" target="_parent">입사자 현황</a>
    <a href="/insa/insa_mg_list.asp" class="btnType01" target="_parent">인사자료미등록현황</a>
    <a href="/insa/insa_report_emp_out.asp" class="btnType01" target="_parent">퇴사자 현황</a>
    <a href="/insa/insa_emp_end_list.asp" class="btnType01" target="_parent">퇴직자조회</a>
    <a href="/insa/insa_emp_infor_main.asp" class="btnType01" target="_parent">인사정보조회</a>
<%If SysAdminYn = "Y" Then%>
    <a href="insa_grade_count.asp" class="btnType01" target="_parent"><span style='color:red;'>직급별 현황</span></a>
    <a href="insa_disabled_list.asp" class="btnType01" target="_parent"><span style='color:red;'>장애인 현황</span></a>
    <a href="insa_report_change.asp" class="btnType01" target="_parent"><span style='color:red;'>인원변동 현황</span></a>
    <a href="insa_report_gunsok.asp" class="btnType01" target="_parent"><span style='color:red;'>근속 현황</span></a>
    <a href="insa_age_count_org.asp" class="btnType01" target="_parent"><span style='color:red;'>연령별 분포</span></a>
    <a href="insa_academic_count_org.asp" class="btnType01" target="_parent"><span style='color:red;'>학력별 분포</span></a>
    <a href="insa_area_count_org.asp" class="btnType01" target="_parent"><span style='color:red;'>지역별 분포</span></a>
    <a href="insa_month_count_mg.asp" class="btnType01" target="_parent"><span style='color:red;'>월별인원</span></a>
    <a href="insa_report_grade.asp" class="btnType01" target="_parent"><span style='color:red;'>승진자 현황</span></a>
<%End If%>
</div>