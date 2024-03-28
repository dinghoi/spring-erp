				<%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				met_grade = request.cookies("nkpmg_user")("coo_met_grade")
				%>  
                <div class="btnRight">
					<a href="met_goods_code_mg.asp" class="btnType01">제.상품 코드</a>
                    <a href="met_stock_code_mg.asp" class="btnType01">창고 코드</a>
                  <% '  <a href="met_stock_code_org.asp" class="btnType01">창고현황(조직)</a> %>
                    <a href="met_stock_code_emp_list.asp" class="btnType01">개인창고 현황</a>
                    <a href="met_control_code_mg.asp" class="btnType01">기타코드관리</a>
                  <% if in_empno = "900002" then	%>						
                        <a href="met_stock_data_check.asp" class="btnType01">조직창고등록</a></li>  
                        <a href="met_stock_data_emp_check.asp" class="btnType01">개인창고등록</a></li>  
                  <% end if	%>  
                  <% if met_grade = "0" then	%>  
                    <a href="#" onClick="pop_Window('met_user_met_grade.asp?u_type=<%="U"%>','met_user_met_grade_pop','scrollbars=yes,width=500,height=350')" class="btnType01">자재관리 권한부여</a>
                  <% end if	%>
				</div>
