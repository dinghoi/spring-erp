				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
				%>  
					<a href="met_stock_move_reg_mg.asp" class="btnType01">â����� �Ƿ�</a>
                    <a href="met_move_chulgo_ing.asp" class="btnType01">â���̵� ���</a>
                    <a href="met_move_chulgo_list.asp" class="btnType01">â����� ��Ȳ</a>
                    <a href="met_move_stin_ing.asp" class="btnType01">â���̵� �԰�</a>
                    <a href="met_move_stin_list.asp" class="btnType01">â���԰� ��Ȳ</a>
                    <a href="met_move_undeliver_list.asp" class="btnType01">â������ ��Ȳ</a>
                    <a href="met_move_stin_not_enter_list.asp" class="btnType01">â����԰� ��Ȳ</a>
                    
             <% '       <a href="meterials_stock_move_mg.asp" class="btnType01">â����� �Ƿ�old</a>
                '    <a href="meterials_chulgo_move.asp" class="btnType01">â���̵� ���old</a> %>
				</div>
