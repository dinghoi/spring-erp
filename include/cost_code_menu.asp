                <div class="btnRight">
                    <a href="cost_grade_mg.asp" class="btnType01">��� ���� ����</a>
                <%
                    '�ֱ漺, ������, ����ȣ, ����ȣ
                    If user_id = "100031" Or user_id = "100359" Or user_id = "100953" Or user_id = "102592" Then
                %>
                    <a href="cost_code_mg.asp" class="btnType01">��뱸���ڵ�</a>
                    <a href="overtime_code_mg.asp" class="btnType01">��Ư�ټ������</a>
                    <a href="insure_per_mg.asp" class="btnType01">4�뺸���������</a>
                <%	End If	%>
                    <a href="trade_cost_mod_mg.asp" class="btnType01">�ŷ�ó����(���)</a>
                <%
                    '�ֱ漺, ������, ����ȣ, ������, ����ȣ
                    If user_id = "100031" Or user_id = "100359" Or user_id = "100953" Or user_id = "101880" Or user_id = "102592" Then
                %>
                    <a href="as_unitprice_mg.asp" class="btnType01">ASǥ�شܰ�</a>
                <%	End If	%>
                </div>
