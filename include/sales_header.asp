<!--#include virtual="/include/google_analytics.asp" -->
<div id="header">
	<h1><!--<img src="/image/com_logo.jpg" alt="홈페이지" width="116" height="30"/>-->
		<img src="/images/logo_long_ko_main.png" alt="K-ONE" width="120" height="40"/>
	</h1>
	<h2>
		<div style="margin:6px;">
			<!--<img src="/image/sales_title.gif" alt="영업 관리 시스템" width="225" height="25"/>-->
			<img src="/image/sales_title.gif" alt="영업 관리 시스템" width="174" height="22"/>
		</div>
	</h2>
	<div class="login">
		<p>
		<strong><%=user_name%>&nbsp;<%=user_grade%>님</strong> 안녕하세요.
		</p>
	</div>
	<div id="gnb">
		<ul>
			<li class="dep1"><a href="/sales/sales_report.asp">매출 전표 관리</a></li>
			<li class="dep1"><a href="/sales_unpaid_mg.asp">미수금 관리</a></li>
			<li class="dep1">
		<%If sales_grade < "2" Or (bonbu = "ITO 사업본부" And position = "사업부장") Or (bonbu = "ITO 사업본부" And position = "본부장") Then	%>
			<a href="/sales/saupbu_profit_loss_total.asp">손익현황</a>
		<%Else	%>
			<a>손익현황</a>
		<%End If%>
			</li>
			<li class="dep1">
		<%If sales_grade = "0" Then	%>
			<a href="/sales_goods_code_mg.asp">코드 관리</a>
		<%Else	%>
			<a>코드 관리</a>
		<%End If%>
			</li>
		</ul>
	</div>
</div>
