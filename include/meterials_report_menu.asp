				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				met_grade = request.cookies("nkpmg_user")("coo_insa_grade")
				%>  
					<a href="met_stock_pummok_inout01.asp" class="btnType01">품목/기간별 입출고현황</a>
                <% 
                    '<a href="met_pummok_inout_mg.asp" class="btnType01">기간별 입출고현황</a>
                %>
                    <a href="met_goods_pum_jaego_mg.asp" class="btnType01">품목별 재고</a>
                    <a href="met_stock_pum_jaego_mg.asp" class="btnType01">창고별 재고</a>
                    <a href="met_goods_pum_jaego_report.asp" class="btnType01">재고(금액순)현황</a>
                    <a href="met_stock_subul_mg.asp" class="btnType01">수불현황</a>
                    <a href="met_stock_jaego_year_trans1.asp" class="btnType01">전기이월</a>
				<% if in_empno = "900002" then	%>		                    
                    <a href="met_stock_chulgo_amt_data.asp" class="btnType01">출고금액</a>
                <% end if	%>                               
                <% '   <a href="meterials_system_popup.asp" class="btnType01">재고실사</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">입고조정</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">출고조정</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">불용재고현황</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">불용재고폐기</a> %>
				</div>
