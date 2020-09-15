<Query Kind="Program">
  <NuGetReference>Microsoft.Graph</NuGetReference>
  <Namespace>Microsoft.Graph</Namespace>
  <Namespace>Newtonsoft.Json.Linq</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>

async Task Main()
{
	var client = V1.GetConfidentialClient();

	await Temp(client);

	//await DownloadDoc(client);
	//await GetUser(client);
	//await GetUsers(client);
	//await CreateMailfolders(client);
	//await CalendarView(client);
	//await CreateGroup(client);
}

static async Task Temp(Microsoft.Graph.GraphServiceClient client)
{
		//var result = await client.Users["adelev@M365x462896.onmicrosoft.com"].Onenote.Pages["1-65b7cbcdd6104f6caa997ca5ede84206!91-76e4ca31-3239-4a7f-a0ef-c61e60ad92ab"].ParentNotebook.Request().GetAsync();
		//var result = await client.Users["adelev@M365x462896.onmicrosoft.com"].Onenote.Pages["1-65b7cbcdd6104f6caa997ca5ede84206!91-76e4ca31-3239-4a7f-a0ef-c61e60ad92ab"].Request().GetAsync();
		//var result2 = await client.Users["adelev@M365x462896.onmicrosoft.com"].Onenote.Pages["1-65b7cbcdd6104f6caa997ca5ede84206!91-76e4ca31-3239-4a7f-a0ef-c61e60ad92ab"].Request().Expand(x => x.ParentSection).GetAsync();
		//var result = await client.Users["adelev@M365x462896.onmicrosoft.com"].Drive.Root.Children.Request().Expand(i => i.Analytics).GetAsync();
		//var result = await client.Users["adelev@M365x462896.onmicrosoft.com"].Drive.Root.Children.Request().GetAsync();
		//var result = await client.Users["adelev@M365x462896.onmicrosoft.com"].Drive.Items["01KA5JMEA7GNMTADXTINDZ7OW7ROGFDU7N"].Request().Expand(i => i.Analytics).GetAsync();
		//var result = await client.Users.Request().Expand("approleassignments").Select("id,mail,displayname,userPrincipalName,MobilePhone,Department,OfficeLocation,UserType,DeletedDateTime,createddatetime").GetAsync();
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

static async Task CalendarView(Microsoft.Graph.GraphServiceClient client)
{
	var queryOptions = new List<QueryOption>()
	{
		new QueryOption("startDateTime", "2020-01-01T19:00:00-08:00Z"),
		new QueryOption("endDateTime", "2020-01-07T19:00:00-08:00Z")
	};

	var calendarView = await client.Users["adelev@M365x462896.onmicrosoft.com"].CalendarView
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