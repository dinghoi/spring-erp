<%
Dim at_name, at_empno, at_position

at_name = request.cookies("nkpmg_user")("coo_user_name")
at_empno = request.cookies("nkpmg_user")("coo_user_id")
at_position = request.cookies("nkpmg_user")("coo_position")
%>
<div class="btnRight">
	<a href="/insa/insa_appoint_mg.asp" class="btnType01">�λ�߷�</a>

	<!--�̻�� �޴� �ӽ� �ּ� ó��[����ȣ_20210402]-->
	<!--<a href="insa_app_bok_mg.asp" class="btnType01">�����߷�</a>
	<a href="insa_appoint_company.asp" class="btnType01">�迭�����߷�</a>-->

	<a href="/insa/insa_report_appoint.asp" class="btnType01">�λ�߷���Ȳ</a>
 <% '   <a href="insa_year_imcome_agree_mg.asp" class="btnType01">�����ٷΰ�ൿ����Ȳ</a> %>

 <!--�̻�� �޴� �ӽ� �ּ� ó��[����ȣ_20210402]-->
 <%' if at_empno = "900002" Or at_empno = "102592" then %>
		<!--<a href="insa_appoint_return.asp" class="btnType01">���Ի��� �̰�</a>-->
 <%' end if %>
</div>
