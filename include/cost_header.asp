<!--#include virtual="/include/google_analytics.asp" -->
			<div id="header">
				<h1><!--<img src="/image/com_logo.jpg" alt="Ȩ������" width="116" height="30"/>-->
					<img src="/images/logo_long_ko_main.png" alt="K-ONE" width="120" height="40"/>
				</h1>
				<h2>
					<div style="margin:6px;">
						<!--<img src="/image/cost_title.gif" alt="��� ���� �ý���" width="189" height="25"/>-->
						<img src="/image/cost_title.gif" alt="��� ���� �ý���" width="166" height="22"/>
					</div>
				</h2>
				<div class="login">
					<p>
                    <strong><%=user_name%>&nbsp;<%=user_grade%>��</strong> �ȳ��ϼ���
					</p>
				</div>
                <div id="gnb">
                    <ul>
                        <li class="dep1"><a href="/person_cost_report.asp">����ϰ���</a></li>
                        <li class="dep1"><a href="/emp_cost_mg.asp">�����Ȳ����</a></li>
                        <li class="dep1">
                        <%
                        '�ֱ漺, ������, ������, ����ȣ, ����ȣ
                        If cost_grade = "0" AND (user_id = "100031" Or user_id = "100359" Or user_id = "101880" Or user_id = "102592" Or user_id = "100953") Then
                        %>
                            <a href="/cost_grade_mg.asp">����ڵ����</a>
                        <%Else	%>
                            <a>����ڵ����</a>
                        <%End If	%>
						</li>
                    </ul>
                </div>
			</div>
