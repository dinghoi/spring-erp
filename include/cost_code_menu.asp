                <div class="btnRight">
                    <a href="cost_grade_mg.asp" class="btnType01">비용 권한 관리</a>
                <%
                    '최길성, 박정신, 하장호, 허정호
                    If user_id = "100031" Or user_id = "100359" Or user_id = "100953" Or user_id = "102592" Then
                %>
                    <a href="cost_code_mg.asp" class="btnType01">비용구분코드</a>
                    <a href="overtime_code_mg.asp" class="btnType01">야특근수당관리</a>
                    <a href="insure_per_mg.asp" class="btnType01">4대보험요율관리</a>
                <%	End If	%>
                    <a href="trade_cost_mod_mg.asp" class="btnType01">거래처변경(비용)</a>
                <%
                    '최길성, 박정신, 하장호, 강경진, 허정호
                    If user_id = "100031" Or user_id = "100359" Or user_id = "100953" Or user_id = "101880" Or user_id = "102592" Then
                %>
                    <a href="as_unitprice_mg.asp" class="btnType01">AS표준단가</a>
                <%	End If	%>
                </div>
