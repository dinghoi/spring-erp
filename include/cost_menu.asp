<div class="btnRight">
	<a href="/person_cost_report.asp" class="btnType01">���κ������</a>
	<a href="/general_cost_mg.asp" class="btnType01">�Ϲݰ��</a>
<%If cost_grade="0" Or cost_grade="3" Then %>
	<a href="/others_cost_mg.asp" class="btnType01">������</a>
<%End If%>
<%If cost_grade="6" Or cost_grade="5" Or cost_grade<"3" Then%>
	<a href="/overtime_mg.asp" class="btnType01">��Ư��</a>
<%End If%>
<%If cost_grade="6" Or cost_grade="5" Or cost_grade<"3" Then%>
	<a href="/transit_cost_mg.asp" class="btnType01">�����</a>
	<!--<a href="/cost/transit_cost_mg.asp" class="btnType01">�����</a>-->
<%End If%>
	<a href="/person_card_mg.asp" class="btnType01">���κ�ī�峻��</a>
<%If cost_grade<"6" Then %>
	<a href="/alba_cost_mg.asp" class="btnType01">�Ƹ�����Ʈ</a>
	<!--<a href="/tax_esero_in_mg.asp" class="btnType01">E���θ��Լ��ݰ�꼭</a>-->
	<a href="/cost/tax_esero_in_mg.asp" class="btnType01">E���θ��Լ��ݰ�꼭</a>
<%If account_grade = "0" Then%>
	<a href="/cost/tax_esero_upload.asp" class="btnType01">E���κ���ϰ����ε�</a>
<%End If%>
	<a href="/cost/tax_bill_in_mg.asp" class="btnType01">���Լ��ݰ�꼭</a>
	<a href="/tax_bill_manual_mg.asp" class="btnType01">���̼��ݰ�꼭</a>
<%End If%>
<%If cost_grade="0" Then %>
	<a href="/cost/depreciation_cost_mg.asp" class="btnType01">�󰢺�</a>
<%End If%>
</div>
