<!--#include virtual="/include/google_analytics.asp" -->
<div id="header">
	<h1><!--<img src="/image/com_logo.jpg" alt="Ȩ������" width="116" height="30"/>-->
		<img src="/images/logo_long_ko_main.png" alt="K-ONE" width="120" height="40"/>
	</h1>
	<h2>
		<div style="margin:6px;">
			<!--<img src="/image/sales_title.gif" alt="���� ���� �ý���" width="225" height="25"/>-->
			<img src="/image/sales_title.gif" alt="���� ���� �ý���" width="174" height="22"/>
		</div>
	</h2>
	<div class="login">
		<p>
		<strong><%=user_name%>&nbsp;<%=user_grade%>��</strong> �ȳ��ϼ���.
		</p>
	</div>
	<div id="gnb">
		<ul>
			<li class="dep1"><a href="/sales/sales_report.asp">���� ��ǥ ����</a></li>
			<li class="dep1"><a href="/sales_unpaid_mg.asp">�̼��� ����</a></li>
			<li class="dep1">
		<%If sales_grade < "2" Or (bonbu = "ITO �������" And position = "�������") Or (bonbu = "ITO �������" And position = "������") Then	%>
			<a href="/sales/saupbu_profit_loss_total.asp">������Ȳ</a>
		<%Else	%>
			<a>������Ȳ</a>
		<%End If%>
			</li>
			<li class="dep1">
		<%If sales_grade = "0" Then	%>
			<a href="/sales_goods_code_mg.asp">�ڵ� ����</a>
		<%Else	%>
			<a>�ڵ� ����</a>
		<%End If%>
			</li>
		</ul>
	</div>
</div>
