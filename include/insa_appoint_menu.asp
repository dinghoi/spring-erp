<%
Dim at_name, at_empno, at_position

at_name = request.cookies("nkpmg_user")("coo_user_name")
at_empno = request.cookies("nkpmg_user")("coo_user_id")
at_position = request.cookies("nkpmg_user")("coo_position")
%>
<div class="btnRight">
	<a href="/insa/insa_appoint_mg.asp" class="btnType01">인사발령</a>

	<!--미사용 메뉴 임시 주석 처리[허정호_20210402]-->
	<!--<a href="insa_app_bok_mg.asp" class="btnType01">복직발령</a>
	<a href="insa_appoint_company.asp" class="btnType01">계열전적발령</a>-->

	<a href="/insa/insa_report_appoint.asp" class="btnType01">인사발령현황</a>
 <% '   <a href="insa_year_imcome_agree_mg.asp" class="btnType01">연봉근로계약동의현황</a> %>

 <!--미사용 메뉴 임시 주석 처리[허정호_20210402]-->
 <%' if at_empno = "900002" Or at_empno = "102592" then %>
		<!--<a href="insa_appoint_return.asp" class="btnType01">재입사자 이관</a>-->
 <%' end if %>
</div>
