				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				met_grade = request.cookies("nkpmg_user")("coo_met_grade")
				%>  
					<% '<a href="met_stock_in_mg.asp" class="btnType01">�����԰���</a> %>
               <% if met_grade = "0" then	%>		                    
                    <a href="#" onClick="pop_Window('met_stock_in_add01.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&u_type=<%=""%>','met_buy_popup','scrollbars=yes,width=1230,height=650')" class="btnType01">�԰���</a>
                    <a href="met_stock_in_report01.asp" class="btnType01">�԰���Ȳ</a>
                    <a href="met_stock_in_report02.asp" class="btnType01">ǰ���԰���Ȳ</a>
               <%	 else	%>
                    <a class="btnType01">�԰���</a>
                    <a class="btnType01">�԰���Ȳ</a>
                    <a class="btnType01">ǰ���԰���Ȳ</a>
               <% end if	%>                                        

                <% '    <a href="meterials_system_popup.asp" class="btnType01">�ŷ�ó��������Ȳ</a>
                   ' <a href="meterials_system_popup.asp" class="btnType01">ǰ�񺰱�����Ȳ</a> %>
				</div>
