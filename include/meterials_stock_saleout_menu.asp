				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				met_grade = request.cookies("nkpmg_user")("coo_met_grade")
				%>  
                    <a href="met_stock_in_sale_report01.asp" class="btnType01">����ǰ�� �԰���Ȳ</a>
                    <a href="met_stock_out_sale_mg.asp" class="btnType01">���� ���� ���</a>
                    <a href="met_stock_out_sale_list.asp" class="btnType01">���� ���� �����Ȳ</a>
				</div>
