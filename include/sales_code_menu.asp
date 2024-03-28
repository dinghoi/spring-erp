<div class="btnRight">
	<a href="/sales_goods_code_mg.asp" class="btnType01">영업품목코드</a>
	<a href="/trade_mod_mg.asp" class="btnType01">거래처변경관리</a>

	<%
	'시스템 관리자 권한만 노출'
	If user_id = "100359" Or user_id = "102592" Then
	%>
	<a href="/sales/sys_log_list.asp" class="btnType01">시스템 로그</a>
	<%
	End If
	%>
</div>
