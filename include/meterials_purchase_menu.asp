				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				%>  
				<% if in_empno = "900002" then	%>						
                       <a href="#" onClick="pop_Window('met_buy_add.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&u_type=<%=""%>','met_buy_popup','scrollbars=yes,width=1230,height=650')" class="btnType01">구매품의등록</a>
                       <a href="met_buy_ing_mg.asp" class="btnType01">구매진행관리</a>
                       <a href="met_buy_order_mg.asp" class="btnType01">발주등록</a>
                       <a href="met_sales_order_mg.asp" class="btnType01">영업발주등록</a>
                       <a href="met_buy_order_ing_mg.asp" class="btnType01">발주진행관리</a>
				<%	 else	%>
                       <a class="btnType01">구매품의등록</a>
                       <a class="btnType01">구매진행관리</a>
                       <a class="btnType01">발주등록</a>
                       <a class="btnType01">영업발주등록</a>
                       <a class="btnType01">발주진행관리</a>
                <% end if	%>                          

                    <% '<a href="meterials_control_mg.asp" class="btnType01">구매품의등록</a>
					   '<a href="meterials_system_popup.asp" class="btnType01">발주서조회</a>
					   '<a href="met_buy_request_mg.asp" class="btnType01">구매요청(접수)현황</a>
                       '<a href="meterials_system_popup.asp" class="btnType01">견적</a> 
					   '<a href="meterials_system_popup.asp" class="btnType01">계약</a> %>
				</div>
