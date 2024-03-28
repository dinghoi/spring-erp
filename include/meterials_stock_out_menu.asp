				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				met_grade = request.cookies("nkpmg_user")("coo_met_grade")
				%>  
				<% if met_grade = "0" then	%>					
                    <a href="#" onClick="pop_Window('met_chulgo_cust_add01.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&rele_id=<%=rele_id%>&u_type=<%=""%>','met_chulgo_reg_pop','scrollbars=yes,width=1230,height=650')" class="btnType01">출고 등록</a>
                   
                    <a href="met_stock_out_reg_ing01.asp" class="btnType01">출고진행관리</a>
                    <a href="met_stock_out_list.asp" class="btnType01">출고현황</a>
				<%	 else	%>                    
                    <a class="btnType01">출고 등록</a>
                    <a class="btnType01">출고진행관리</a>
                    <a class="btnType01">출고현황</a>
                <% end if	%>            
                <% '    <a href="met_chulgo_cust_list.asp" class="btnType01">인수증 테스트 ing</a>
                    ' <a href="meterials_system_popup.asp" class="btnType01">고객반납입고 ing</a>
                    '   <a href="meterials_system_popup.asp" class="btnType01">고객미반납현황</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">고객미출고현황</a>
                    '<a href="met_chulgo_cust_list.asp" class="btnType01">고객출고현황</a> %>
				</div>
