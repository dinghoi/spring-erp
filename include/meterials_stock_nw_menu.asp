				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				met_grade = request.cookies("nkpmg_user")("coo_insa_grade")
				%>
                    <a href="met_import_nwin_report01.asp" class="btnType01">외자입고현황</a>
                    <a href="met_stock_nwin_report01.asp" class="btnType01">N/W입고현황</a>
                    <a href="met_import_out_sale_mg.asp" class="btnType01">N/W 고객사 출고</a>
                    <a href="met_import_out_sale_list.asp" class="btnType01">N/W 고객사 출고현황</a>
                    <a href="met_import_serial_list.asp" class="btnType01">Serial관리대장</a>
				<% '외자 자재관리등 개발 이전 프로그램으로 참조용임..추후 지워버리면 됨
				   if in_empno = "100001" Or in_empno = "102592" Then %>
				    <a href="#" onClick="pop_Window('met_import_nwin_add01.asp?view_condi=<%=view_condi%>&goods_type=<%="상품"%>&u_type=<%=""%>','met_import_nwin_add01_popup','scrollbars=yes,width=1300,height=650')" class="btnType01">외자입고등록</a>
					<a href="#" onClick="pop_Window('met_stock_nwin_add01.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&u_type=<%=""%>','met_stock_nwin_add01_popup','scrollbars=yes,width=1230,height=650')" class="btnType01">N/W입고등록</a>
					<a href="#" onClick="pop_Window('met_chulgo_nwcust_add01.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&rele_id=<%=rele_id%>&u_type=<%=""%>','met_nwchulgo_reg_pop','scrollbars=yes,width=1230,height=650')" class="btnType01">N/W출고등록</a>
                    <a href="met_stock_nwout_reg_ing01.asp" class="btnType01">N/W출고현황</a>
				<%  end if %>
				</div>
