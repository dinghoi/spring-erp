				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				met_grade = request.cookies("nkpmg_user")("coo_insa_grade")
				%>
                    <a href="met_import_nwin_report01.asp" class="btnType01">�����԰���Ȳ</a>
                    <a href="met_stock_nwin_report01.asp" class="btnType01">N/W�԰���Ȳ</a>
                    <a href="met_import_out_sale_mg.asp" class="btnType01">N/W ���� ���</a>
                    <a href="met_import_out_sale_list.asp" class="btnType01">N/W ���� �����Ȳ</a>
                    <a href="met_import_serial_list.asp" class="btnType01">Serial��������</a>
				<% '���� ��������� ���� ���� ���α׷����� ��������..���� ���������� ��
				   if in_empno = "100001" Or in_empno = "102592" Then %>
				    <a href="#" onClick="pop_Window('met_import_nwin_add01.asp?view_condi=<%=view_condi%>&goods_type=<%="��ǰ"%>&u_type=<%=""%>','met_import_nwin_add01_popup','scrollbars=yes,width=1300,height=650')" class="btnType01">�����԰���</a>
					<a href="#" onClick="pop_Window('met_stock_nwin_add01.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&u_type=<%=""%>','met_stock_nwin_add01_popup','scrollbars=yes,width=1230,height=650')" class="btnType01">N/W�԰���</a>
					<a href="#" onClick="pop_Window('met_chulgo_nwcust_add01.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&rele_id=<%=rele_id%>&u_type=<%=""%>','met_nwchulgo_reg_pop','scrollbars=yes,width=1230,height=650')" class="btnType01">N/W�����</a>
                    <a href="met_stock_nwout_reg_ing01.asp" class="btnType01">N/W�����Ȳ</a>
				<%  end if %>
				</div>
