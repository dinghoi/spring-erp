				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				%>  
				<% if in_empno = "900002" then	%>						
                       <a href="#" onClick="pop_Window('met_buy_add.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&u_type=<%=""%>','met_buy_popup','scrollbars=yes,width=1230,height=650')" class="btnType01">����ǰ�ǵ��</a>
                       <a href="met_buy_ing_mg.asp" class="btnType01">�����������</a>
                       <a href="met_buy_order_mg.asp" class="btnType01">���ֵ��</a>
                       <a href="met_sales_order_mg.asp" class="btnType01">�������ֵ��</a>
                       <a href="met_buy_order_ing_mg.asp" class="btnType01">�����������</a>
				<%	 else	%>
                       <a class="btnType01">����ǰ�ǵ��</a>
                       <a class="btnType01">�����������</a>
                       <a class="btnType01">���ֵ��</a>
                       <a class="btnType01">�������ֵ��</a>
                       <a class="btnType01">�����������</a>
                <% end if	%>                          

                    <% '<a href="meterials_control_mg.asp" class="btnType01">����ǰ�ǵ��</a>
					   '<a href="meterials_system_popup.asp" class="btnType01">���ּ���ȸ</a>
					   '<a href="met_buy_request_mg.asp" class="btnType01">���ſ�û(����)��Ȳ</a>
                       '<a href="meterials_system_popup.asp" class="btnType01">����</a> 
					   '<a href="meterials_system_popup.asp" class="btnType01">���</a> %>
				</div>
