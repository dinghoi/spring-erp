<!--#include virtual="/include/google_analytics.asp" -->
<div id="header">
	<h1><!--<img src="/image/com_logo.jpg" alt="Ȩ������" width="116" height="30"/>-->
		<img src="/images/logo_long_ko_main.png" alt="K-ONE" width="120" height="40"/>
	</h1>
	<h2>
		<div style="margin:6px;">
			<!--<img src="/image/pay_title.gif" alt="�λ�޿� �ý���" width="198" height="25"/>-->
			<img src="/image/pay_title.gif" alt="�λ�޿� �ý���" width="174" height="22"/>
		</div>
	</h2>
	<div class="login">
		<p>
		<strong><%=request.cookies("nkpmg_user")("coo_user_name")%>&nbsp;<%=request.cookies("nkpmg_user")("coo_user_grade")%></strong>�� �ȳ��ϼ���.
		</p>
	</div>
	<div id="gnb">
		<ul>
			<li class="dep1"><a href="/pay/insa_pay_mg.asp">��(��)������</a></li>
			<li class="dep1"><a href="/pay/insa_pay_insurance_mg.asp">�޿����ʼ���</a></li>
			<li class="dep1"><a href="/pay/insa_pay_code_mg.asp">�޿��ڵ����</a></li>
		   	<%
		   	If SysAdminYn = "Y" Then
			%>
				<li class="dep1"><a href="/insa_pay_empout_annual.asp"><span style="color:red;">�����޿�</span></a></li>
				<li class="dep1"><a href="/insa_pay_alba_mg.asp"><span style="color:red;">����ҵ�޿�����</span></a></li>
				<li class="dep1"><a href="/insa_pay_year_income_mg.asp"><span style="color:red;">����/�������װ���</span></a></li>
				<li class="dep1"><a href="insa_social_mg.asp"><span style="color:red;">��ȸ����</span></a></li>
				<li class="dep1"><a href="/insa_pay_tax_mg.asp"><span style="color:red;">�����Ű����</span></a></li>
				<li class="dep1"><a href="insa_system_popup.asp"><span style="color:red;">�����������</span></a></li>
				<li class="dep1"><a href="/insa_pay_pay_report_mg.asp"><span style="color:red;">��Ȳ/���</span></a></li>
				<li class="dep1"><a href="insa_pay_yeartax2_mg.asp"><span style="color:red;">��������</span></a></li>
			<%
		   	End If
		   	%>
		</ul>
	</div>
</div>
