				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				met_grade = request.cookies("nkpmg_user")("coo_met_grade")
				%>  
					<% '<a href="met_stock_in_mg.asp" class="btnType01">구매입고등록</a> %>
               <% if met_grade = "0" then	%>		                    
                    <a href="#" onClick="pop_Window('met_stock_in_add01.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&u_type=<%=""%>','met_buy_popup','scrollbars=yes,width=1230,height=650')" class="btnType01">입고등록</a>
                    <a href="met_stock_in_report01.asp" class="btnType01">입고현황</a>
                    <a href="met_stock_in_report02.asp" class="btnType01">품목별입고현황</a>
               <%	 else	%>
                    <a class="btnType01">입고등록</a>
                    <a class="btnType01">입고현황</a>
                    <a class="btnType01">품목별입고현황</a>
               <% end if	%>                                        

                <% '    <a href="meterials_system_popup.asp" class="btnType01">거래처별구매현황</a>
                   ' <a href="meterials_system_popup.asp" class="btnType01">품목별구매현황</a> %>
				</div>
