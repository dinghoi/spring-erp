				<div class="btnRight">
					<a href="/emp_cost_mg.asp" class="btnType01">개인별 비용</a>
				<% If cost_grade < "3" Or position = "팀장" Or position = "사업부장" Or position = "본부장" Then  %>
					<a href="/org_person_cal_report.asp" class="btnType01">조직별 개인별 비용</a>
		        <% End If	%>
				<% If cost_grade = "0" Or user_id = "100031" Or user_id = "100178" Or user_id = "100029" Then  %>
					<a href="/reside_person_cal_report.asp" class="btnType01">상주처별 개인별 비용</a>
				<% End If  %>
				<% If user_id = "100359" Or user_id = "102592" Then %>
					<a href="/cost/cost_emp_master_month_mg.asp" class="btnType01">직원 월별 현황</a>
		        <% End If	%>
				<% If user_id = "100952" Or user_id = "100359" Or user_id = "101100" Or user_id = "102305" Or user_id = "102306" Or user_id = "102592" Then %>
					<a href="/cost/cost_end_mg.asp" class="btnType01">비용마감관리</a>
					<!--<a href="/cost_end/cost_end_mg.asp" class="btnType01">비용마감관리</a>-->
                <% End If	%>
				<% If user_id = "900001" Or user_id = "100359"  Or user_id = "100952" Or user_id = "101100" Or user_id = "102305" Or user_id = "102306"  Or user_id = "102592" Then %>
					<a href="/cost/cost_end_condi_cancel.asp" class="btnType01">비용마감일괄취소</a>
					<!--<a href="/cost_end/cost_end_condi_cancel.asp" class="btnType01">비용마감일괄취소</a>-->
                <% End If	%>
				<% If cost_grade < "3" Or position = "사업부장" Or position = "본부장" Then  %>
					<a href="/saupbu_cost_report.asp" class="btnType01">사업부별 비용</a>
                <% End If	%>
				<% If cost_grade < "2" Or position = "본부장" Or user_id = "100031" Then  %>
					<a href="/cost_bonbu_approval_mg.asp" class="btnType01">본부 비용 승인</a>
                <% End If	%>
				<% If cost_grade = "0" Then  %>
					<a href="/total_cost_report.asp" class="btnType01">회사 전체 비용</a>
                <% End If	%>
				</div>
