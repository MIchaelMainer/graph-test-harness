<Query Kind="Program">
  <NuGetReference>Microsoft.Graph</NuGetReference>
  <NuGetReference Prerelease="true">Microsoft.Graph.Auth</NuGetReference>
  <Namespace>Microsoft.Graph</Namespace>
  <Namespace>Microsoft.Graph.Auth</Namespace>
  <Namespace>Microsoft.Graph.Auth.Extensions</Namespace>
  <Namespace>Microsoft.Graph.CallRecords</Namespace>
  <Namespace>Microsoft.Graph.Core.Requests</Namespace>
  <Namespace>Microsoft.Graph.Extensions</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>Newtonsoft.Json.Bson</Namespace>
  <Namespace>Newtonsoft.Json.Converters</Namespace>
  <Namespace>Newtonsoft.Json.Linq</Namespace>
  <Namespace>Newtonsoft.Json.Schema</Namespace>
  <Namespace>Newtonsoft.Json.Serialization</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Runtime.Serialization.Formatters</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
  <DisableMyExtensions>true</DisableMyExtensions>
</Query>

async Task Main()
{
	var chambele_client_confidential = Chambele.V1.GetConfidentialGraphClient();

	await DriveItemSearch(chambele_client_confidential);
	//await Temp(chambele_client_confidential);
}

static async Task Temp(Microsoft.Graph.GraphServiceClient client)
{

	//await client.AppCatalogs.TeamsApps[""].Request().CreateAsync(
	//var result = await client.Users["adelev@M365x462896.onmicrosoft.com"].Onenote.Pages["1-65b7cbcdd6104f6caa997ca5ede84206!91-76e4ca31-3239-4a7f-a0ef-c61e60ad92ab"].ParentNotebook.Request().GetAsync();
	//var result = await client.Users["adelev@M365x462896.onmicrosoft.com"].Onenote.Pages["1-65b7cbcdd6104f6caa997ca5ede84206!91-76e4ca31-3239-4a7f-a0ef-c61e60ad92ab"].Request().GetAsync();
	//var result2 = await client.Users["adelev@M365x462896.onmicrosoft.com"].Onenote.Pages["1-65b7cbcdd6104f6caa997ca5ede84206!91-76e4ca31-3239-4a7f-a0ef-c61e60ad92ab"].Request().Expand(x => x.ParentSection).GetAsync();
	//var result = await client.Users["adelev@M365x462896.onmicrosoft.com"].Drive.Root.Children.Request().Expand(i => i.Analytics).GetAsync();
	//var result = await client.Users["adelev@M365x462896.onmicrosoft.com"].Drive.Root.Children.Request().GetAsync();
	//var result = await client.Users["adelev@M365x462896.onmicrosoft.com"].Drive.Items["01KA5JMEA7GNMTADXTINDZ7OW7ROGFDU7N"].Request().Expand(i => i.Analytics).GetAsync();
	//var result = await client.Users.Request().Expand("approleassignments").Select("id,mail,displayname,userPrincipalName,MobilePhone,Department,OfficeLocation,UserType,DeletedDateTime,createddatetime").GetAsync();
}

static async Task HttpTemp(GraphServiceClient client)
{
	var allApps = await client.DeviceAppManagement.MobileApps.Request().Filter("isOf('microsoft.graph.win32LobApp')").GetAsync();
	var allCategories = await client.DeviceAppManagement.MobileAppCategories.Request().GetAsync();
	var hrm = client.DeviceAppManagement.MobileApps["mobileAppId"].Categories.References.Request().GetHttpRequestMessage();
	hrm.Method = HttpMethod.Post;
	hrm.Content = new StringContent($"{{\"@odata.id\": \"https://graph.microsoft.com/v1.0/deviceAppManagement/mobileAppCategories/{allCategories[2]}\"}}");
	var response = await client.HttpProvider.SendAsync(hrm);

}

static async Task GetJoinWebUrl(GraphServiceClient client)
{
	string body = string.Format("joinweburl%20eq%20'{0}'", Chambele.V1.JoinWebUrl);

	var meetingsPage = await client.Users[Chambele.V1.user].OnlineMeetings.Request().Filter(body).GetAsync();
	meetingsPage.Dump(nameof(meetingsPage));
}

static async Task SearchDrive(Microsoft.Graph.GraphServiceClient client)
{
	var searchResults = await client.Users[Chambele.V1.user].Drives["b!PjyRhY4Wj0GXmwrlvqF0qz5yxyhyQaNNmv-RlWgxtlyGoRo1wOogSb-XMgXCnxJO"].Search("folder").Request().GetAsync();
	searchResults.Dump(nameof(searchResults));
}

static async Task Group_GetNonExistentGroup(Microsoft.Graph.GraphServiceClient client)
{
	var group = await client.Groups["676aff12-2525-4463-97da-5550ff425d05"].Request().GetAsync();
	group.Dump();
}

static async Task ValidatePrimitiveMethod(Microsoft.Graph.GraphServiceClient client)
{
	await client.Users[Chambele.V1.user].Drive.Items["01DVF26FGQUWUQODOYUZGZ7CJ5YDPDI6J4"].Workbook.Worksheets["sheet1"].Charts.Count().Request().GetAsync();
}

static async Task Composable(Microsoft.Graph.GraphServiceClient client)
{
	//await client.Drive.Items["fileId"].Workbook.Worksheets["worksheetId"].Range("A1").
}

static async Task TestBatch()
{
	var httpClient = Chambele.V1.GetConfidentialHttpClient();
	// Make dummy request to get token into cache
	HttpRequestMessage httpRequestMessage_dummy = new HttpRequestMessage(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/users/{Chambele.V1.user}");
	Console.WriteLine("Get token");
	var stopwatchToken = Stopwatch.StartNew();
	var result = await httpClient.SendAsync(httpRequestMessage_dummy);
	stopwatchToken.Stop();
	Console.WriteLine(stopwatchToken.Elapsed.TotalMilliseconds.ToString());

	// Manual
	HttpRequestMessage httpRequestMessage_m1 = new HttpRequestMessage(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/users/{Chambele.V1.user}");
	HttpRequestMessage httpRequestMessage_m2 = new HttpRequestMessage(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/users/{Chambele.V1.user}/calendars");
	HttpRequestMessage httpRequestMessage_m3 = new HttpRequestMessage(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/users/{Chambele.V1.user}/messages/{Chambele.V1.messageWithAttachment_Id}/attachments/{Chambele.V1.attachmentId}/$value");
	HttpRequestMessage httpRequestMessage_m4 = new HttpRequestMessage(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/users/{Chambele.V1.user}/messages/{Chambele.V1.messageWithAttachment_Id}/$value");
	HttpRequestMessage httpRequestMessage_m5 = new HttpRequestMessage(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/users/{Chambele.V1.user}/drive/root/search(q='finance')?select=name,id,webUrl");
	
	var stopwatch1 = Stopwatch.StartNew();
	await httpClient.SendAsync(httpRequestMessage_m1);
	await httpClient.SendAsync(httpRequestMessage_m2);
	//await httpClient.SendAsync(httpRequestMessage_m3);
	//await httpClient.SendAsync(httpRequestMessage_m4);
	//await httpClient.SendAsync(httpRequestMessage_m5);
	stopwatch1.Stop();

	// Batch
	HttpRequestMessage httpRequestMessage_b1 = new HttpRequestMessage(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/users/{Chambele.V1.user}");
	HttpRequestMessage httpRequestMessage_b2 = new HttpRequestMessage(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/users/{Chambele.V1.user}/calendars");
	HttpRequestMessage httpRequestMessage_b3 = new HttpRequestMessage(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/users/{Chambele.V1.user}/messages/{Chambele.V1.messageWithAttachment_Id}/attachments/{Chambele.V1.attachmentId}/$value");
	HttpRequestMessage httpRequestMessage_b4 = new HttpRequestMessage(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/users/{Chambele.V1.user}/messages/{Chambele.V1.messageWithAttachment_Id}/$value");
	HttpRequestMessage httpRequestMessage_b5 = new HttpRequestMessage(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/users/{Chambele.V1.user}/drive/root/search(q='finance')?select=name,id,webUrl");

	// Add batch request steps to BatchRequestContent.
	BatchRequestContent batchRequestContent = new BatchRequestContent();
	batchRequestContent.AddBatchRequestStep(httpRequestMessage_b1);
	batchRequestContent.AddBatchRequestStep(httpRequestMessage_b2);
	//batchRequestContent.AddBatchRequestStep(httpRequestMessage_b3);
	//batchRequestContent.AddBatchRequestStep(httpRequestMessage_b4);
	//batchRequestContent.AddBatchRequestStep(httpRequestMessage_b5);

	
	var batchStopwatch = Stopwatch.StartNew();
	// Send batch request with BatchRequestContent. 
	await httpClient.PostAsync("https://graph.microsoft.com/v1.0/$batch", batchRequestContent);
	batchStopwatch.Stop();

	Console.WriteLine("Batch");
	Console.WriteLine(batchStopwatch.Elapsed.TotalMilliseconds.ToString());	
	Console.WriteLine("Manual");
	Console.WriteLine(stopwatch1.Elapsed.TotalMilliseconds.ToString());
}


static async Task UseChambele(Microsoft.Graph.GraphServiceClient client)
{
	
	
	await CreateAndUpdateEvent(client, Chambele.V1.user);
	await CreateAndUpdateEventUTC(client, Chambele.V1.user);
}

static async Task DriveItemSearch(Microsoft.Graph.GraphServiceClient client)
{
	var result = client.Users[Chambele.V1.user].Drive.Root.Search().Request().GetAsync();
}

static async Task CreateAndUpdateEvent(Microsoft.Graph.GraphServiceClient client, string user)
{
	var @event = new Event
	{
		Subject = "Test subject - Created with PST",
		Body = new ItemBody { Content = "Test body content" },
		Start = new DateTimeTimeZone { DateTime = "2020-09-16T18:00:00.0000000", TimeZone = "Pacific Standard Time" },
		End = new DateTimeTimeZone { DateTime = "2020-09-16T18:30:00.0000000", TimeZone = "Pacific Standard Time" }
	};

	var result = await client.Users[user].Events.Request().AddAsync(@event);
	result.Dump();

	var patchEventObject = new Event
	{
		Start = new DateTimeTimeZone { DateTime = "2020-09-16T19:00:00.0000000", TimeZone = "Pacific Standard Time" },
		End = new DateTimeTimeZone { DateTime = "2020-09-16T19:30:00.0000000", TimeZone = "Pacific Standard Time" }
	};
		
	var result2 = await client.Users[user].Events[result.Id].Request().UpdateAsync(patchEventObject);
	result2.Dump();
}

static async Task CreateAndUpdateEventUTC(Microsoft.Graph.GraphServiceClient client, string user)
{
	var @event = new Event
	{
		Subject = "Test subject - Created with UTC",
		Body = new ItemBody { Content = "Test body content" },
		Start = new DateTimeTimeZone { DateTime = "2020-09-17T01:00:00Z", TimeZone = "UTC" },
		End = new DateTimeTimeZone { DateTime = "2020-09-17T01:30:00Z", TimeZone = "UTC" }
	};

	var result = await client.Users[user].Events.Request().AddAsync(@event);
	result.Dump();

	var patchEventObject = new Event
	{   // UTC is 8 hours ahead of PST

		Start = new DateTimeTimeZone { DateTime = "2020-09-17T02:00:00.0000000", TimeZone = "UTC" },
		End = new DateTimeTimeZone { DateTime = "2020-09-17T02:30:00.0000000", TimeZone = "UTC" }

		// This format works as well
		//Start = new DateTimeTimeZone { DateTime = "2020-09-17T02:00:00Z", TimeZone = "UTC" },
		//End = new DateTimeTimeZone { DateTime = "2020-09-17T02:30:00Z", TimeZone = "UTC" }
		
	};

	var result2 = await client.Users[user].Events[result.Id].Request().UpdateAsync(patchEventObject);
	result2.Dump();
}

static async Task CreateGroup(Microsoft.Graph.GraphServiceClient client)
{
	var group = new Microsoft.Graph.Group()
	{
		Description = "Group with designated owner and members",
		DisplayName = "Operations group2",
		GroupTypes = new List<String>()
		{
			"Unified"
		},
		MailEnabled = true,
		MailNickname = "operations20192",
		SecurityEnabled = false,
		AdditionalData = new Dictionary<string, object>() {
			{"owners@odata.bind", new JArray() {
				"https://graph.microsoft.com/v1.0/users/adelev@M365x462896.onmicrosoft.com"
			}},
			{"members@odata.bind", new JArray() {
				"https://graph.microsoft.com/v1.0/users/adelev@M365x462896.onmicrosoft.com"
			}}
		}	
	};

	await client.Groups
	.Request()
	.AddAsync(group);
}

static async Task GetMessagesWithOptions(Microsoft.Graph.GraphServiceClient client)
{
	var queryOptions = new List<QueryOption>()
	{
		new QueryOption("select", "subject"),
		new QueryOption("$count", "true")
	};

	var messages = await client.Users[Chambele.V1.user].Messages
   								.Request(queryOptions)
								.GetAsync();

	messages.Dump();
}

static async Task GetHiddenFolders(Microsoft.Graph.GraphServiceClient client)
{
	var queryOptions = new List<QueryOption>()
	{
		new QueryOption("includeHiddenFolders", "true")
	};

	var hiddenFolders = await client.Users[Chambele.V1.user]
									.MailFolders
   									.Request(queryOptions)
									.GetAsync();

	hiddenFolders.Dump();
}

static async Task CalendarView(Microsoft.Graph.GraphServiceClient client)
{
	var queryOptions = new List<QueryOption>()
	{
		new QueryOption("startDateTime", "2020-01-01T19:00:00-08:00Z"),
		new QueryOption("endDateTime", "2020-01-07T19:00:00-08:00Z")
	};

	var calendarView = await client.Users[Chambele.V1.user].CalendarView
		.Request(queryOptions)
		.GetAsync();
		
	calendarView.Dump();		
}

static async Task CreateEvent(Microsoft.Graph.GraphServiceClient client)
{
	var appointment = new Event();
	appointment.Subject = "MyEvent from lib";
	appointment.Start = new DateTimeTimeZone();
	appointment.Start.DateTime = "2020-07-22T23:00:00";
	appointment.Start.TimeZone = "Pacific Standard Time";
	appointment.End = new DateTimeTimeZone();
	appointment.End.DateTime = "2020-07-23T00:00:00";
	appointment.End.TimeZone = "Pacific Standard Time";

	var result = client.Users["adelev@M365x462896.onmicrosoft.com"].Events.Request().AddAsync(appointment);

	//var result = client.Users["adelev@M365x462896.onmicrosoft.com"].Events["AAMkADc4MGE5NGY5LTJmY2EtNGIzMy1hNDI1LTQwYWNmNmJjYmEyZABGAAAAAAAydp9QX0V8TI8yN6kVwIE7BwDzysquKi1lSbmXhC_kv0IxAAAAAAENAADzysquKi1lSbmXhC_kv0IxAAKnIbQPAAA="].Request().GetAsync();
	result.Dump();
}

static async Task CreateMailfolders(GraphServiceClient client)
{

	for (int n = 0; n < 11; n++)
	{

		var folder = new MailFolder()
		{
			DisplayName = $"folder{n.ToString()}"
		};

		var result = await client.Users["adelev@M365x462896.onmicrosoft.com"].MailFolders["inbox"].ChildFolders.Request().AddAsync(folder);
	}
}

static async Task GetUsers(GraphServiceClient client)
{
	IEnumerable<User> users = await client.Users.Request().GetAsync();
	var listOfUsers = users.ToList();
	
	listOfUsers.Dump();
}

static async Task SendEmailWithExpiration(GraphServiceClient client)
{
	var message = new Message();
	//client.HttpProvider.SendAsync(httpRequestMessage)
	
	var expiryPropPage = new MessageSingleValueExtendedPropertiesCollectionPage();
	var expiryProp = new SingleValueLegacyExtendedProperty();
	expiryProp.Id = "SystemTime 0x0015";
	
	
	// GET /me/mailFolders/{id}/messages/{id}?$expand=singleValueExtendedProperties($filter=id eq '{id_value}')
	// ?$expand=singleValueExtendedProperties($filter=id eq 'SystemTime 0x0015')
	
	// https://graph.microsoft.com/v1.0/me/messages/AAMkADEzOTExYjJkLTYxZDAtNDgxOC04YzQyLTU0OGY1Yzc3ZGY0MwBGAAAAAADhS2QUsLGoTbY_lhGktZkcBwCfcL8jd7fqRJCX-i4H_BeNAAAD7kV3AACYnYF59gJhQaoeGgGqm4QrAAd4whPUAAA=
	
}

static async Task DownloadDoc(GraphServiceClient client)
{
	var result = await client.Users["adelev@M365x462896.onmicrosoft.com"].Drive.Items["01KA5JMEGYLK7ZXO2KLZCJSDS6DDWCZVTR"].Content.Request().GetAsync();
	result.Dump();
}

static async Task<string> GetUser(GraphServiceClient client)
{
	var userPage = await client.Users["adelev@M365x462896.onmicrosoft.com"].Request().GetAsync();
	return userPage.Id;
}


static async Task GetAttachmentAsMIME(GraphServiceClient client)
{
	var a =  client.Me.Messages["msg-id"].Attachments["att-id"];

	var ar = new AttachmentRequest("", null, null);
	ar.AppendSegmentToRequestUrl("$value");

//	public System.Threading.Tasks.Task<string> GetAsMIMEAsync(CancellationToken cancellationToken)
//	{
//		this.ContentType = "text/plain";
//		this.Method = "GET";
//		this.AppendSegmentToRequestUrl("$value");
//		return await this.SendAsync<string>(null, cancellationToken).ConfigureAwait(false);
//	}

}

// Define other methods and classes here
static async Task DeltaQuery(GraphServiceClient client)
{
	string deltaToken = "";
	IDriveItemDeltaRequest nextRequest = client.Sites.Root.Drive.Root.Delta(null).Request().Top(1);


	while (nextRequest != null)
	{
		HeaderOption hierarchicalSharingOption = new HeaderOption("prefer", "hierarchicalsharing");
		HeaderOption showSharingChangesOption = new HeaderOption("prefer", "deltashowsharingchanges");
		nextRequest.Headers.Add(hierarchicalSharingOption);
		nextRequest.Headers.Add(showSharingChangesOption);

		IDriveItemDeltaCollectionPage deltaResults = await nextRequest.GetAsync();
		foreach (DriveItem driveItem in deltaResults)
		{
			Console.WriteLine(driveItem.Name);
		}

		nextRequest = deltaResults.NextPageRequest;
		if (nextRequest != null)
		{
			
			Console.WriteLine("Press enter to get next page of results...");
			Console.ReadLine();
		}
		else
		{
			deltaToken = deltaResults.AdditionalData["@odata.deltaLink"].ToString();
		}
	}

}

// Get clients setup to use the chambele tenant. Only uses the v1 client.
namespace Chambele
{
	public static class V1
	{
		public static string user = "michael@chambele.onmicrosoft.com";
		public static string messageWithAttachment_Id = "AAMkADE4NmViZTM1LTE5NWYtNDM5NS05ZWNiLTAzOGMzMTJlZDZhOQBGAAAAAABQxBLf_AP_TIifsgqw7XDwBwDbMXxdZjcWTIGiiYxXPUznAAAAAAEMAADbMXxdZjcWTIGiiYxXPUznAAAtZdglAAA=";
		public static string attachmentId = "AAMkADE4NmViZTM1LTE5NWYtNDM5NS05ZWNiLTAzOGMzMTJlZDZhOQBGAAAAAABQxBLf_AP_TIifsgqw7XDwBwDbMXxdZjcWTIGiiYxXPUznAAAAAAEMAADbMXxdZjcWTIGiiYxXPUznAAAtZdglAAABEgAQALmGlhZxBNRKhrRzIWv3B4I=";
		public const string JoinWebUrl = "https://teams.microsoft.com/l/meetup-join/19%3ameeting_ZDU5YmM4Y2YtYTgxNC00NDdlLTk5NDEtYzIzNmM4OTYxMmQy%40thread.v2/0?context=%7b%22Tid%22%3a%22fe639da9-7429-46bf-9a68-40cfc8229591%22%2c%22Oid%22%3a%225cae63f5-0216-4766-8ca3-84d5e288e796%22%7d";
		
		private static IAuthenticationProvider GetClientCredentialProvider()
		{
			var authClient = Microsoft.Identity.Client.ConfidentialClientApplicationBuilder
							   .Create(Util.GetPassword("chambele_clientId"))
							   .WithTenantId(Util.GetPassword("chambele_tenantId"))
							   .WithClientSecret(Util.GetPassword("chambele_clientsecret"))
							   .Build();
		
			return new ClientCredentialProvider(authClient);
		}
		
		public static Microsoft.Graph.GraphServiceClient GetConfidentialGraphClient()
		{
			return new GraphServiceClient(GetClientCredentialProvider());
		}
		
		public static HttpClient GetConfidentialHttpClient()
		{
			return GraphClientFactory.Create(GetClientCredentialProvider());
		}
	}
}