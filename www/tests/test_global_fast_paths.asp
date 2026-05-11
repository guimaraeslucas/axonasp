<%
Dim g_i, g_sum, g_limit, g_count
g_sum = 0
g_limit = 3
For g_i = 1 To g_limit
	g_sum = g_sum + g_i
Next
g_count = 1
g_count = g_count + 1
g_count = g_count - 1
Response.Write g_sum & "|" & g_i & "|" & g_count
%>