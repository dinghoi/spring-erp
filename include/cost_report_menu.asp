				<div class="btnRight">
					<a href="/emp_cost_mg.asp" class="btnType01">���κ� ���</a>
				<% If cost_grade < "3" Or position = "����" Or position = "�������" Or position = "������" Then  %>
					<a href="/org_person_cal_report.asp" class="btnType01">������ ���κ� ���</a>
		        <% End If	%>
				<% If cost_grade = "0" Or user_id = "100031" Or user_id = "100178" Or user_id = "100029" Then  %>
					<a href="/reside_person_cal_report.asp" class="btnType01">����ó�� ���κ� ���</a>
				<% End If  %>
				<% If user_id = "100359" Or user_id = "102592" Then %>
					<a href="/cost/cost_emp_master_month_mg.asp" class="btnType01">���� ���� ��Ȳ</a>
		        <% End If	%>
				<% If user_id = "100952" Or user_id = "100359" Or user_id = "101100" Or user_id = "102305" Or user_id = "102306" Or user_id = "102592" Then %>
					<a href="/cost/cost_end_mg.asp" class="btnType01">��븶������</a>
					<!--<a href="/cost_end/cost_end_mg.asp" class="btnType01">��븶������</a>-->
                <% End If	%>
				<% If user_id = "900001" Or user_id = "100359"  Or user_id = "100952" Or user_id = "101100" Or user_id = "102305" Or user_id = "102306"  Or user_id = "102592" Then %>
					<a href="/cost/cost_end_condi_cancel.asp" class="btnType01">��븶���ϰ����</a>
					<!--<a href="/cost_end/cost_end_condi_cancel.asp" class="btnType01">��븶���ϰ����</a>-->
                <% End If	%>
				<% If cost_grade < "3" Or position = "�������" Or position = "������" Then  %>
					<a href="/saupbu_cost_report.asp" class="btnType01">����κ� ���</a>
                <% End If	%>
				<% If cost_grade < "2" Or position = "������" Or user_id = "100031" Then  %>
					<a href="/cost_bonbu_approval_mg.asp" class="btnType01">���� ��� ����</a>
                <% End If	%>
				<% If cost_grade = "0" Then  %>
					<a href="/total_cost_report.asp" class="btnType01">ȸ�� ��ü ���</a>
                <% End If	%>
				</div>
