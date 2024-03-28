				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
				%>  
					<a href="met_stock_move_reg_mg.asp" class="btnType01">창고출고 의뢰</a>
                    <a href="met_move_chulgo_ing.asp" class="btnType01">창고이동 출고</a>
                    <a href="met_move_chulgo_list.asp" class="btnType01">창고출고 현황</a>
                    <a href="met_move_stin_ing.asp" class="btnType01">창고이동 입고</a>
                    <a href="met_move_stin_list.asp" class="btnType01">창고입고 현황</a>
                    <a href="met_move_undeliver_list.asp" class="btnType01">창고미출고 현황</a>
                    <a href="met_move_stin_not_enter_list.asp" class="btnType01">창고미입고 현황</a>
                    
             <% '       <a href="meterials_stock_move_mg.asp" class="btnType01">창고출고 의뢰old</a>
                '    <a href="meterials_chulgo_move.asp" class="btnType01">창고이동 출고old</a> %>
				</div>
