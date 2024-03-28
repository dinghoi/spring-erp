<div class="btnRight">
	<a href="/as_list_ce.asp" class="btnType01">나의 A/S 현황</a>
<%If mg_group = "2" Then %>
	<a href="/as_acpt_reg_2.asp" class="btnType01">A/S 접수 등록</a>
<%Else %>
	<a href="/as_acpt_reg.asp" class="btnType01">A/S 접수 등록</a>
<%End If %>
<%
If user_id = "102592" Then
%>
	<a href="/as_acpt_reg_new.asp" class="btnType01">A/S 접수 등록(개발중)</a>
<%End If	%>

	<a href="/into_list.asp" class="btnType01">입고관리</a>

<%If reside_company = "국세청" Then	%>
	<a href="/as_list_reside.asp" class="btnType01">A/S 총괄 현황(상주처)</a>
<%Else	%>
	<a href="/service/as_list.asp" class="btnType01">A/S 총괄 현황</a>
<%End If	%>

	<a href="/att_list.asp" class="btnType01">설치/공사 첨부관리</a>
	<a href="/company_form_mg.asp" class="btnType01">회사별 양식 관리</a>

	<a href="/service/as_acpt_statics_list.asp" class="btnType01">월별 A/S 현황</a>
	<a href="/service/as_acpt_statics_up.asp" class="btnType01">월별 A/S 현황 업로드</a>
</div>
