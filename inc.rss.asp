<script language="vbscript" type="text/vbscript" runat="server">

	
	'
	' Simple Rss Feed Reader 1.2
	'
	' This source snippet provides easy access to any Rss Feed by wrapping the
	' response Xml into class instances with properties.
	' 
	' See http://web.resource.org/rss/1.0/spec for Rss Specification
	' 
	' The script *requires* the use of IIS5+ installations where the 
	' "Microsoft.XMLDOM" COM object is available.
	' 
	' <author name="Peter Theill" />
	' <date release="2004-04-21T22:57:00Z" />
	'
	
	
	'
	' Retrieves RSS feed from specified URL
	'
	Function GetRss(url)
		
		' create new CRss object and load specified url
		Dim rss: Set rss = new CRss
		rss.Url = url
			
		Set GetRss = rss
		
	End Function
	
	
	Class CRss
		Private url_
		Private channel_
		
		Private Sub Class_Initialize()
			Set channel_ = Nothing
		End Sub
		
		public Property Get Url
			Url = url_
		End Property
		
		Public Property Let Url(v)
			url_ = v
			
			Dim domXml: Set domXml = Server.CreateObject("Microsoft.XMLDOM")
			domXml.Async = False
			domXml.SetProperty "ServerHTTPRequest", True
			domXml.ResolveExternals = True
			domXml.ValidateOnParse = True
			domXml.Load(url_)
			
			If (domXml.parseError.errorCode = 0) Then
				Dim rootNode: Set rootNode = domXml.documentElement
				If (NOT IsObject(rootNode)) Then
					Err.Raise vbObjectError + 3, "Xml Data", "No Root element found", "", 0
					Exit Property
				End If
				
				Dim channelNode: Set channelNode = rootNode.selectSingleNode("channel")
				If (NOT IsObject(channelNode)) Then
					Err.Raise vbObjectError + 4, "Xml Data", "No 'channel' element found", "", 0
					Exit Property
				End If
				
				' read channel info
				Set channel_ = New CRssChannel
				channel_.Title = channelNode.selectSingleNode("title").Text
				channel_.Description = channelNode.selectSingleNode("description").Text
				channel_.Link = channelNode.selectSingleNode("link").Text
				
				' read items within channel
				Dim objLinks: Set objLinks = rootNode.getElementsByTagName("item")
				If (IsObject(objLinks)) Then
					Dim objChild
					For Each objChild in objLinks
						Dim ri: Set ri = New CRssItem
						ri.Title = objChild.selectSingleNode("title").Text
						ri.Description = objChild.selectSingleNode("description").Text
						ri.Link = objChild.selectSingleNode("link").Text
						
						channel_.AddItem(ri)
					Next
				End If
				
				' release used resources
				Set rootNode = Nothing
				Set channelNode = Nothing
				Set objLinks = Nothing
			Else
				Err.Raise vbObjectError + 2, "Xml Data", _
					"Unable to parse Xml: " & _
					domXml.parseError.reason, _
					"", _
					0
			End If
			
			Set domXml = Nothing
		End Property
		
		Public Property Get Channel
			Set Channel = channel_
		End Property
		
	End Class ' // > Class Rss
	
	
	Class CRssChannel 
		Private items_
		
		Public Title
		Public Description
		Public Link
		Public Image
		
		Private Sub Class_Initialize()
			Set items_ = Server.CreateObject("Scripting.Dictionary")
		End Sub
		
		Private Sub Class_Terminate()
			Set items_ = Nothing
		End Sub
		
		Public Sub AddItem(v)
			items_.Add items_.Count, v
		End Sub
		
		Public Property Get Items
			Items = items_.Items
		End Property
		
		Public Property Let Items(v)
			Set items_ = v
		End Property
	
	End Class ' // > Class CRssChannel
	
	
	Class CRssItem
		Public Title
		Public Description
		Public Link
	
	End Class ' // > Class CRssItem
	
</script>