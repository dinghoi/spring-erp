				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				met_grade = request.cookies("nkpmg_user")("coo_met_grade")
				%>  
                    <a href="met_move_stin_list01.asp" class="btnType01">�԰���Ȳ</a>
                    <a href="met_stock_move_ce_mg.asp" class="btnType01">CE���(â��)</a>
                    <a href="met_stock_move_mg.asp" class="btnType01">CE���(â��)��Ȳ</a>
                    <a href="met_stock_out_ce_mg.asp" class="btnType01">��(��)���</a>
                    <a href="met_stock_out_mg.asp" class="btnType01">��(��)�����Ȳ</a>
				</div>
