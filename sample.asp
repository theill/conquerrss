<%
	
	Option Explicit
	
	'
	'
	'
	
%>
<!-- #include file="inc.rss.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
	"http://www.w3.org/TR/html4/loose.dtd">
<html>

<head>
	<meta name="description" content="free asp source for a fast online chat client">
	<meta name="keywords" content="ConquerChat, asp, free source, online chat, chat online, chat, theill, peter, fontlister, conquerware, conquerchat, fillout manager, svendsk, swatchitime, reklamer, meninger, java, delphi, about">
	<title>ConquerRSS | Sample</title>
	<link rel="stylesheet" type="text/css" href="../../css/theill.css">
	<link rel="stylesheet" type="text/css" href="../../css/asp.css">
	<link rel="stylesheet" type="text/css" href="css/rss.css">
</head>

<body class="site">
<div align="center"><center>

<table border="0" width="590" cellspacing="0" cellpadding="0">
  <tr align="top">
    <td bgcolor="#ffffff" valign="top"><table border="0" cellspacing="0" cellpadding="4">
      <tr>
        <td valign="top">
        <!--webbot bot="Include" U-Include="../../__menu.asp" TAG="BODY" startspan --><strong>[__menu.asp]</strong><!--webbot bot="Include" endspan i-checksum="4782" --></td>
        <td width="78%" valign="top"><div align="center"><center><table border="0" width="446" cellspacing="0" cellpadding="0">
          <tr>
            <td width="100%" valign="middle" height="4"><img src="../../images/dot.gif" width="20" height="4" border="0"></td>
          </tr>
          <tr>
            <td width="446" valign="middle" height="58"><img border="0" src="../../images/asp_topbar.gif" alt="asp @ theill.com" width="446" height="58"></td>
          </tr>
          <tr>
            <td width="100%" valign="top" align="justify">
              <p class="Navigation"><b>navigation:</b> <a href="/">root</a>&nbsp; 
              / <a href="/asp.asp">asp</a> / <a href="/asp/conquerrss.asp">Conquer<b>RSS</b></a></p>
              <h2 class="subSectionHeader">Conquer<b>RSS</b></h2>
              <h4 class="subSectionHeader">Get RSS feeds delivered directly into your web site</h4>
              
<%

	Dim rssFeedUrl: rssFeedUrl = Request("rss_feed_url")
	If (rssFeedUrl = "") Then
		rssFeedUrl = "http://www.gotdotnet.com/team/dbox/rss.aspx"
	End If

%>
              <center>
              <form>
              	<input type="text" name="rss_feed_url" size="48" value="<%= Server.HtmlEncode(rssFeedUrl) %>" />
              	<input type="submit" value="Get Feed!" />
              </form>
              </center>
              <p align="left">
<%
	
'	On Error Resume Next
	Dim rss: Set rss = GetRss(rssFeedUrl)
	If (IsObject(rss)) Then
		Dim chn: Set chn = rss.Channel
		
		Response.Write("<table cellspacing='0' cellpadding='4' class='rss'>")
		Response.Write("<thead class='rss'>")
		Response.Write("<th class='rss'>")
		Response.Write("<a href='" & chn.Link & "' class='rss' title='" & chn.Description & "'>" & chn.Title & "</a> RSS feed")
		Response.Write("</th>")
		Response.Write("</thead>")
		
		Response.Write("<tbody>")
		
		Dim lnk
		For Each lnk in chn.Items
			Response.Write("<tr>")
			Response.Write("<td>")
			Response.Write("<a href='" & lnk.Link & "' class='rss' target='_blank'>" & lnk.Title & "</a>")
			Response.Write("<div class='rss'>" & lnk.Description & "</div>")
			Response.Write("</td>")
			Response.Write("</tr>")
		Next
		Response.Write("</tbody>")
		
		Response.Write("</table>")
		
		' release used resources
		Set rss = Nothing
	Else
		Response.Write( _
			"Could not read RSS:<br />" & _
			"&nbsp;Url: " & rssFeedUrl & "<br />" & _
			"&nbsp;Code: " & Err.Number & "<br />" & _
			"&nbsp;Description: " & Err.Description)
	End If
	
%>
              </p>
							
							Here are some feeds to try:
							<ul>
								<li>http://www.simtel.com/rss.php</li>
								<li>http://msdn.microsoft.com/rss.xml</li>
								<li>http://msdn.microsoft.com/vcsharp/rss.xml</li>
								<li>http://msdn.microsoft.com/webservices/rss.xml</li>
								<li>http://www.gotdotnet.com/team/dbox/rss.aspx</li>
							</ul>

              </td>
          </tr>
        </table>
        </div></td>
      </tr>
    </table>
    </td>
  </tr>
  <tr align="top">
    <td valign="top" align="right">
	<!--webbot bot="Include" U-Include="../../__footer.asp" TAG="BODY" startspan --><strong>[__footer.asp]</strong><!--webbot bot="Include" endspan i-checksum="9964" --></td>
  </tr>
</table>
</div>
</body>
</html>