
 ConquerRSS README
 Copyright (c) 2004 Peter Theill, theill.com
 
 -----------------------------------------------------------------------------
 Introduction
 -----------------------------------------------------------------------------
 This script is a simple RSS Feed Reader.
 
 -----------------------------------------------------------------------------
 Usage
 -----------------------------------------------------------------------------
 Paste the following code to a new ASP file such as "testing.asp".


	<!-- #include file="inc.rss.asp" -->
	Dim rss: Set rss = GetRss("http://msdn.microsoft.com/rss.xml")
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
		Response.Write("Could not read RSS from " & rssFeedUrl)
	End If


 -----------------------------------------------------------------------------
 Support
 -----------------------------------------------------------------------------
 Direct mails will *not* be answered, sorry. All support *must* go through the
 forum at http://www.theill.com/forum.asp
 
 The script requires "Microsoft.XMLDOM" which is installed on IIS5 and up.
 