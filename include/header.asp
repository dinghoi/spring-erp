<!--#include virtual="/include/google_analytics.asp" -->
			<div id="header">
				<h1><!--<img src="/image/com_logo.jpg" alt="Ȩ������" width="116" height="30"/>-->
					<img src="/images/logo_long_ko_main.png" alt="K-ONE" width="120" height="40"/>
				</h1>
				<h2>
					<div style="margin:6px;">
						<!--<img src="/image/main_title.gif" alt="A/S ���� �ý���" width="225" height="25"/>-->
						<img src="/image/main_title.gif" alt="A/S ���� �ý���" width="198" height="22"/>
					</div>
				</h2>
				<div class="login">
					<p>
                    <strong><%=user_name%>&nbsp;<%=user_grade%>��</strong> �ȳ��ϼ���.
					</p>
				</div>
                <div id="gnb">
                    <ul>
                        <li class="dep1"><a href="/as_list_ce.asp">A/S ����</a></li>
                        <li class="dep1">
					<%If c_grade < "1" Then	%>
                        <a href="/large_data_up.asp">�ٷ�ó��</a>
					<%Else	%>
                        <a>�ٷ�ó��</a>
                    <%End If%>
                        </li>
                        <li class="dep1">
                    <%If c_grade < "2" Then	%>
                        <a href="/waiting.asp?pg_name=<%="day_sum.asp"%>">�Ѱ���Ȳ</a>
                    <%Else	%>
                        <a>�Ѱ���Ȳ</a>
                    <%End If	%>
                        </li>
                        <li class="dep1">
                    <%If c_grade < "2" Then	%>
                        <a href="/waiting.asp?pg_name=<%="area_term_pro.asp"%>">�����Ȳ</a>
					<%Else%>
                        <a>�����Ȳ</a>
                    <%End If%>
                        </li>
                        <li class="dep1">
                    <%If c_grade < "1" Then	%>
                        <a href="/ce_mg_list.asp">CE ����</a></li>
					<%Else	%>
                        <a>CE ����</a>
					<%End If%>
                        </li>
                        <li class="dep1">
                    <%If c_grade < "1" Then	%>
                        <a href="/etc_code_mg.asp">�ڵ� ����</a>
					<%Else	%>
                        <a>�ڵ� ����</a>
                    <%End If	%>
                        </li>
                    </ul>
                </div>
			</div>
