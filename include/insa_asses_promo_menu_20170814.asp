				<%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				%>  
                <div class="btnRight">
                    <a href="insa_promotion_list.asp" class="btnType01">승진대상자현황</a>
                <%  if in_empno = "100787" or in_empno = "100952" or in_empno = "101086" or in_empno = "101168" or in_empno = "101485" then %>                    
                    <a href="insa_emp_owner_org_list.asp" class="btnType01">상위조직변경</a>
                <%  end if %>

				<% If in_empno="101100" Or in_empno="100952" Or in_empno="101086" Then %>
                    <a href="insa_pay_total_info.asp" class="btnType01" target="_parent">사업부별 인건비조회</a>
				<% End If %>
				</div>
