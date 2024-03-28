<div class="btnRight">
    <a href="/insa/insa_mg.asp" class="btnType01" target="_parent">직원현황</a>
    <a href="/insa/insa_approve_mg.asp" class="btnType01" target="_parent">채용승인</a>
    <a href="/insa/insa_master_modify.asp" class="btnType01" target="_parent">인사기본정보</a>
    <a href="/insa/insa_family_mg.asp" class="btnType01" target="_parent">가족사항</a>
    <a href="/insa/insa_school_mg.asp" class="btnType01" target="_parent">학력사항</a>
    <a href="/insa/insa_career_mg.asp" class="btnType01" target="_parent">경력사항</a>
    <a href="/insa/insa_qual_mg.asp" class="btnType01" target="_parent">자격사항</a>
    <a href="/insa/insa_edu_mg.asp" class="btnType01" target="_parent">교육사항</a>
    <a href="/insa/insa_language_mg.asp" class="btnType01" target="_parent">어학능력</a>
<%If SysAdminYn = "Y" Then%>
    <a href="/insa/insa_emp_master_mg.asp" class="btnType01" target="_parent"><span style='color:red;'>직원별 관리</span></a>
    <a href="insa_emp_add01.asp" class="btnType01"><span style='color:red;'>신규채용등록</span></a>
    <a href="/insa/insa_reward_punish_mg.asp" class="btnType01" target="_parent"><span style='color:red;'>상벌사항</span></a>
    <a href="insa_system_popup.asp" class="btnType01"><span style='color:red;'>신원보증</span></a>
    <a href="/insa_emp_yryc_list.asp" class="btnType01" target="_parent"><span style='color:red;'>근속1년미만</span></a>
    <a href="/insa/insa_emp_org_list.asp" class="btnType01" target="_parent"><span style='color:red;'>조직별직원현황</span></a>
    <a href="/insa_comment_list.asp" class="btnType01" target="_parent"><span style='color:red;'>특이사항</span></a>
    <a href="insa_master_month_mg.asp" class="btnType01" target="_parent"><span style='color:red;'>인사마감</span></a>
<%End If%>
</div>