<!--#include virtual="/include/google_analytics.asp" -->
			<div id="header">
				<h1><img src="/image/com_logo.jpg" alt="Ȩ������" width="116" height="30"/></h1>
				<h2><img src="/image/main_title.gif" alt="A/S ���� �ý���" width="198" height="25"/></h2>
				<div class="login">
					<p>
                    <strong><%=user_name%>&nbsp;<%=user_grade%></strong>�� �ȳ��ϼ���.
					</p>
					<a href="#" onclick="javascript:pop_user_mod();"><img src="/image/user_mod.gif" alt="������������"/></a>
                    <a href="logout.asp"><img src="/image/logout.gif" alt="�α׾ƿ�"/></a>
				</div>	
                <div id="gnb">
                    <ul>
                        <li class="dep1"><a href="as_list_ce_user.asp">A/S ����</a></li>    
                        <li class="dep1">  
						
                    <%	if team <> "���ְ���" then	%>						
                        <a href="waiting.asp?pg_name=<%="day_sum_user.asp"%>">��Ȳ����</a>
					<%	  else	%>
                        <a>��Ȳ����</a>
						<%	end if	%>
                         <li>
						 <%	if user_name <> "�Ե���Ż" then	%>						
                        <a href="waiting.asp?pg_name=<%="nkp_main2.asp"%>">��������</a>
					<%else%>
                        <a>��������</a>
						<%	end if	%>
						</li>
                    </ul>
                </div>
			</div>
