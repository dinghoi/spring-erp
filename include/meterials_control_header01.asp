<!--#include virtual="/include/google_analytics.asp" -->
<!--#include virtual="/common/common.asp" -->
<div id="header">
	<h1><!--<img src="/image/com_logo.jpg" alt="Ȩ������" width="116" height="30"/>-->
		<img src="/images/logo_long_ko_main.png" alt="K-ONE" width="120" height="40"/>
	</h1>
	<h2>
		<div style="margin:6px;">
			<!--<img src="/image/meterials_control_title.gif" alt="��ǰ������� �ý���" width="198" height="25"/>-->
			<img src="/image/meterials_control_title.gif" alt="��ǰ������� �ý���" width="221" height="22"/>
		</div>
	</h2>
	<div class="login">
		<p>
		<strong><%=request.cookies("nkpmg_user")("coo_user_name")%>&nbsp;<%=request.cookies("nkpmg_user")("coo_user_grade")%></strong>�� �ȳ��ϼ���.
		</p>
	</div>
	<%
	in_name = request.cookies("nkpmg_user")("coo_user_name")
	in_empno = request.cookies("nkpmg_user")("coo_user_id")
	position = request.cookies("nkpmg_user")("coo_position")
	met_grade = request.cookies("nkpmg_user")("coo_met_grade")
	%>
	<div id="gnb">
		<ul>
		<% if met_grade = "0" then	%>
			<li class="dep1"><a href="met_stock_in_report01.asp">�԰� ����</a></li>
			<li class="dep1"><a href="met_stock_out_reg_ing01.asp">��� ����</a></li>
		<%	 else	%>
			<li class="dep1"><a>�԰� ����</a>
			<li class="dep1"><a>��� ����</a>
		<% end if	%>
		<%'������, ������, ����ȣ
		'if met_grade = "2" or in_empno = "101100" or in_empno = "100359" or in_empno = "100015" or in_empno = "100442" or in_empno = "100397" Or in_empno = "102592"  then
		If met_grade = "2" Or NwInMenuYn = "Y" Then
		%>
			<li class="dep1"><a href="met_stock_nwin_report01.asp">N/W�����</a></li>
		<%Else %>
			<li class="dep1"><a>N/W�����</a>
		<%End If %>
			<li class="dep1"><a href="met_stock_out_ce_mg.asp">CE�����</a></li>
			<li class="dep1"><a href="met_stock_jaego_mg.asp">��� ����</a></li>
			<li class="dep1">
		<% if met_grade = "0" or met_grade = "2" then	%>
			<a href="met_stock_pum_jaego_mg.asp">��Ȳ �� ���</a></li>
		<%	 else	%>
			<a>��Ȳ �� ���</a>
		<% end if	%>
			<li class="dep1">
		<% if met_grade = "0" or met_grade = "2" then	%>
			<a href="met_goods_code_mg.asp">�ڵ� ����</a></li>
		<%	 else	%>
			<a>�ڵ� ����</a>
		<% end if	%>
			<li class="dep1"><a href="met_stock_out_sale_mg.asp">�������</a></li>
		</ul>
	</div>
</div>