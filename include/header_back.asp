			<div id="header">
				<h1><img src="/image/com_logo.jpg" alt="Ȩ������" width="116" height="30"/></h1>
				<h2><img src="/image/main_title.gif" alt="A/S ���� �ý���" width="225" height="25"/></h2>
				<%
				mi_view = 0
				Sql = "select count(*) from sign_msg where recv_id = '"&user_id&"' and read_yn = 'N'"
				Set rs_mi = Dbconn.Execute (sql)
				mi_view = cint(rs_mi(0))
				rs_mi.close()
				%>
				<div class="login">
					<p>
                    <strong><%=user_name%>&nbsp;<%=user_grade%>��.</strong>
					</p>
					<a href="#" onClick="pop_Window('sign_process_mg.asp','sign_process_mg_pop','scrollbars=yes,width=1250,height=600')"><img src="image/close_icon.gif" width="16" height="13"><%=mi_view%>��&nbsp;</a>
					<a href="#" onclick="javascript:pop_user_mod();"><img src="/image/user_mod.gif" alt="������������"/></a>
                    <a href="logout.asp"><img src="/image/logout.gif" alt="�α׾ƿ�"/></a>
				</div>	
                <div id="gnb">
                    <ul>
                        <li class="dep1"><a href="board_list.asp">A/S ����</a></li>    
                        <li class="dep1">
					<% if c_grade < "1" then	%>						
                        <a href="large_data_up.asp">�ٷ�ó��</a>
					<%	 else	%>
                        <a>�ٷ�ó��</a>
                    <% end if	%>
                        </li>                        
                        <li class="dep1"><a href="overtime_mg.asp">������</a></li>                        
                        <li class="dep1"><a href="cost_list.asp">�����Ȳ</a></li>                        
                        <li class="dep1">
                    <%	if c_grade < "1" then	%>						
                        <a href="waiting.asp?pg_name=<%="day_sum.asp"%>">�Ѱ���Ȳ</a>
					<%	  else	%>
                        <a>�Ѱ���Ȳ</a>
                    <%	end if	%>
                        </li>
                        <li class="dep1">
                    <%	if c_grade < "1" then	%>						
                        <a href="waiting.asp?pg_name=<%="area_term_pro.asp"%>">�����Ȳ</a>
					<%	  else	%>
                        <a>�����Ȳ</a>
                    <%	end if	%>
                        </li>
                        <li class="dep1">
                    <%	if c_grade < "1" then	%>						
                        <a href="ce_mg_list.asp">CE ����</a></li>
					<%	  else	%>
                        <a>����ڰ���</a>
					<%	end if	%>
                        </li>
                        <li class="dep1">
                    <%	if c_grade < "1" then	%>						
                        <a href="etc_code_mg.asp">�ڵ� ����</a>
					<%	  else	%>
                        <a>�ڵ� ����</a>
                    <%	end if	%>
                        </li>
                    </ul>
                </div>
			</div>
