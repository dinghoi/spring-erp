<!--#include virtual="/include/google_analytics.asp" -->
<!--#include virtual="/common/common.asp" -->
<div id="header">
	<h1><!--<img src="/image/com_logo.jpg" alt="홈페이지" width="116" height="30"/>-->
		<img src="/images/logo_long_ko_main.png" alt="K-ONE" width="120" height="40"/>
	</h1>
	<h2>
		<div style="margin:6px;">
			<!--<img src="/image/meterials_control_title.gif" alt="상품자재관리 시스템" width="198" height="25"/>-->
			<img src="/image/meterials_control_title.gif" alt="상품자재관리 시스템" width="221" height="22"/>
		</div>
	</h2>
	<div class="login">
		<p>
		<strong><%=request.cookies("nkpmg_user")("coo_user_name")%>&nbsp;<%=request.cookies("nkpmg_user")("coo_user_grade")%></strong>님 안녕하세요.
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
			<li class="dep1"><a href="met_stock_in_report01.asp">입고 관리</a></li>
			<li class="dep1"><a href="met_stock_out_reg_ing01.asp">출고 관리</a></li>
		<%	 else	%>
			<li class="dep1"><a>입고 관리</a>
			<li class="dep1"><a>출고 관리</a>
		<% end if	%>
		<%'박정신, 전간수, 허정호
		'if met_grade = "2" or in_empno = "101100" or in_empno = "100359" or in_empno = "100015" or in_empno = "100442" or in_empno = "100397" Or in_empno = "102592"  then
		If met_grade = "2" Or NwInMenuYn = "Y" Then
		%>
			<li class="dep1"><a href="met_stock_nwin_report01.asp">N/W입출고</a></li>
		<%Else %>
			<li class="dep1"><a>N/W입출고</a>
		<%End If %>
			<li class="dep1"><a href="met_stock_out_ce_mg.asp">CE입출고</a></li>
			<li class="dep1"><a href="met_stock_jaego_mg.asp">재고 관리</a></li>
			<li class="dep1">
		<% if met_grade = "0" or met_grade = "2" then	%>
			<a href="met_stock_pum_jaego_mg.asp">현황 및 출력</a></li>
		<%	 else	%>
			<a>현황 및 출력</a>
		<% end if	%>
			<li class="dep1">
		<% if met_grade = "0" or met_grade = "2" then	%>
			<a href="met_goods_code_mg.asp">코드 관리</a></li>
		<%	 else	%>
			<a>코드 관리</a>
		<% end if	%>
			<li class="dep1"><a href="met_stock_out_sale_mg.asp">영업출고</a></li>
		</ul>
	</div>
</div>