				<div class="btnRight">
					<a href="person_cost_report.asp" class="btnType01">���κ��������</a>
					<a href="general_cost_mg.asp" class="btnType01">�Ϲݰ��</a>
				<% if cost_grade = "0" or cost_grade = "3" then  %>
					<a href="others_cost_mg.asp" class="btnType01">��������</a>
                <% end if	%>
				<% if cost_grade = "6" or cost_grade = "5" or cost_grade < "3" then  %>
					<a href="overtime_mg.asp" class="btnType01">��Ư��</a>
                <% end if	%>
				<% if cost_grade = "6" or cost_grade = "5" or cost_grade < "3" then  %>
					<a href="transit_cost_mg.asp" class="btnType01">�����</a>
                <% end if	%>
					<a href="person_card_mg.asp" class="btnType01">���κ�ī�峻��</a>
				<% if cost_grade < "6" then  %>
					<a href="tax_esero_in_mg.asp" class="btnType01">E���θ��Լ��ݰ�꼭</a>
					<a href="rent_cost_mg.asp" class="btnType01">������</a>
					<a href="outside_cost_mg.asp" class="btnType01">���ֺ�</a>
					<a href="parcel_cost_mg.asp" class="btnType01">��ݺ�</a>
					<a href="etc_cost_mg.asp" class="btnType01">��������</a>
					<a href="tax_bill_cost_mg.asp" class="btnType01">��꼭�Ϲݺ��</a>
                <% end if	%>
				</div>
