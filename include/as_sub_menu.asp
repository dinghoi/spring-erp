<div class="btnRight">
	<a href="/as_list_ce.asp" class="btnType01">���� A/S ��Ȳ</a>
<%If mg_group = "2" Then %>
	<a href="/as_acpt_reg_2.asp" class="btnType01">A/S ���� ���</a>
<%Else %>
	<a href="/as_acpt_reg.asp" class="btnType01">A/S ���� ���</a>
<%End If %>
<%
If user_id = "102592" Then
%>
	<a href="/as_acpt_reg_new.asp" class="btnType01">A/S ���� ���(������)</a>
<%End If	%>

	<a href="/into_list.asp" class="btnType01">�԰����</a>

<%If reside_company = "����û" Then	%>
	<a href="/as_list_reside.asp" class="btnType01">A/S �Ѱ� ��Ȳ(����ó)</a>
<%Else	%>
	<a href="/service/as_list.asp" class="btnType01">A/S �Ѱ� ��Ȳ</a>
<%End If	%>

	<a href="/att_list.asp" class="btnType01">��ġ/���� ÷�ΰ���</a>
	<a href="/company_form_mg.asp" class="btnType01">ȸ�纰 ��� ����</a>

	<a href="/service/as_acpt_statics_list.asp" class="btnType01">���� A/S ��Ȳ</a>
	<a href="/service/as_acpt_statics_up.asp" class="btnType01">���� A/S ��Ȳ ���ε�</a>
</div>
