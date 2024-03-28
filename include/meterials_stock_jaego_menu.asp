				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
				%>  
                    <a href="met_goods_jaego_mg.asp" class="btnType01">품목/창고별 재고</a>
                    <a href="met_stock_jaego_mg.asp" class="btnType01">창고/품목별 재고</a>
                    <a href="met_goods_jaego_list.asp" class="btnType01">품목별 재고현황</a>
                <% if  in_empno  = "900002" then   %>
                    <a href="met_organiz_jaego_mg.asp" class="btnType01">조직별 재고</a>
                <% end if  %>
                <%  '<a href="met_stock_move_jaego.asp" class="btnType01">이동중 재고</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">적정재고</a>
				    '<a href="meterials_system_popup.asp" class="btnType01">재고실사</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">입고조정</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">출고조정</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">불용재고현황</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">불용재고폐기</a> %>
				</div>
