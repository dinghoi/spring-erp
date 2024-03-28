<!--#include virtual="/include/google_analytics.asp" -->
<div id="header">
	<h1><!--<img src="/image/com_logo.jpg" alt="홈페이지" width="116" height="30"/>-->
		<img src="/images/logo_long_ko_main.png" alt="K-ONE" width="120" height="40"/>
	</h1>
	<h2>
		<div style="margin:6px;">
			<!--<img src="/image/pay_title.gif" alt="인사급여 시스템" width="198" height="25"/>-->
			<img src="/image/pay_title.gif" alt="인사급여 시스템" width="174" height="22"/>
		</div>
	</h2>
	<div class="login">
		<p>
		<strong><%=request.cookies("nkpmg_user")("coo_user_name")%>&nbsp;<%=request.cookies("nkpmg_user")("coo_user_grade")%></strong>님 안녕하세요.
		</p>
	</div>
	<div id="gnb">
		<ul>
			<li class="dep1"><a href="/pay/insa_pay_mg.asp">급(상)여관리</a></li>
			<li class="dep1"><a href="/pay/insa_pay_insurance_mg.asp">급여기초설정</a></li>
			<li class="dep1"><a href="/pay/insa_pay_code_mg.asp">급여코드관리</a></li>
		   	<%
		   	If SysAdminYn = "Y" Then
			%>
				<li class="dep1"><a href="/insa_pay_empout_annual.asp"><span style="color:red;">퇴직급여</span></a></li>
				<li class="dep1"><a href="/insa_pay_alba_mg.asp"><span style="color:red;">사업소득급여관리</span></a></li>
				<li class="dep1"><a href="/insa_pay_year_income_mg.asp"><span style="color:red;">연봉/보수월액관리</span></a></li>
				<li class="dep1"><a href="insa_social_mg.asp"><span style="color:red;">사회보험</span></a></li>
				<li class="dep1"><a href="/insa_pay_tax_mg.asp"><span style="color:red;">세무신고관리</span></a></li>
				<li class="dep1"><a href="insa_system_popup.asp"><span style="color:red;">연말정산관리</span></a></li>
				<li class="dep1"><a href="/insa_pay_pay_report_mg.asp"><span style="color:red;">현황/출력</span></a></li>
				<li class="dep1"><a href="insa_pay_yeartax2_mg.asp"><span style="color:red;">연말정산</span></a></li>
			<%
		   	End If
		   	%>
		</ul>
	</div>
</div>
