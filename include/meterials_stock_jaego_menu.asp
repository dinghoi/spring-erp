				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
				%>  
                    <a href="met_goods_jaego_mg.asp" class="btnType01">ǰ��/â�� ���</a>
                    <a href="met_stock_jaego_mg.asp" class="btnType01">â��/ǰ�� ���</a>
                    <a href="met_goods_jaego_list.asp" class="btnType01">ǰ�� �����Ȳ</a>
                <% if  in_empno  = "900002" then   %>
                    <a href="met_organiz_jaego_mg.asp" class="btnType01">������ ���</a>
                <% end if  %>
                <%  '<a href="met_stock_move_jaego.asp" class="btnType01">�̵��� ���</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">�������</a>
				    '<a href="meterials_system_popup.asp" class="btnType01">���ǻ�</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">�԰�����</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">�������</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">�ҿ������Ȳ</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">�ҿ�������</a> %>
				</div>
