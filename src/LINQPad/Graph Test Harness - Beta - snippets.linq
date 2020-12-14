<Query Kind="Program">
  <NuGetReference Prerelease="true">Microsoft.Graph.Auth</NuGetReference>
  <NuGetReference Version="0.29.0-preview" Prerelease="true">Microsoft.Graph.Beta</NuGetReference>
  <NuGetReference>Microsoft.Identity.Client</NuGetReference>
  <Namespace>Chambele</Namespace>
  <Namespace>Microsoft.Graph</Namespace>
  <Namespace>Microsoft.Graph.Auth</Namespace>
  <Namespace>Microsoft.Graph.Auth.Extensions</Namespace>
  <Namespace>Microsoft.Graph.CallRecords</Namespace>
  <Namespace>Microsoft.Graph.Core.Requests</Namespace>
  <Namespace>Microsoft.Graph.Extensions</Namespace>
  <Namespace>Microsoft.Graph.TermStore</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>Newtonsoft.Json.Bson</Namespace>
  <Namespace>Newtonsoft.Json.Converters</Namespace>
  <Namespace>Newtonsoft.Json.Linq</Namespace>
  <Namespace>Newtonsoft.Json.Schema</Namespace>
  <Namespace>Newtonsoft.Json.Serialization</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Runtime.Serialization.Formatters</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
  <Namespace>Microsoft.Identity.Client</Namespace>
  <Namespace>System.Net.Http.Headers</Namespace>
</Query>

async Task Main()
{
	//var client = Chambele.Beta.GetPublicGraphClient();
	var client = Chambele.Beta.GetConfidentialGraphClient();

	await ExternalItem(client);
}

static async Task ExternalItem(Microsoft.Graph.GraphServiceClient client)
{
	client.Users[userId].Authentication.Methods.Request().GetAsync()
	
	
	await client.External.Connections.Request().GetAsync().Dump();
}

static async Task Temp(Microsoft.Graph.GraphServiceClient client)
{
	//var todoTask = new TodoTask();
	//todoTask.Title = "Subject";
	//todoTask.DueDateTime = new DateTimeTimeZone() { DateTime = DateTime.Now.AddDays(5).ToString() };
	//todoTask.Status = Microsoft.Graph.TaskStatus.NotStarted;
	//todoTask.Importance = Importance.Normal;
	//todoTask.Body = new ItemBody
	//{
	//	Content = "Test",
	//	ContentType = BodyType.Text
	//};
	//todoTask.IsReminderOn = true;
	//todoTask.ReminderDateTime = new DateTimeTimeZone()
	//{
	//	DateTime = DateTime.Now.AddDays(4).ToString()
	//};
	//todoTask.Extensions = new TodoTaskExtensionsCollectionPage();
	//todoTask.Extensions.Add(new OpenTypeExtension
	//{
	//	ExtensionName = "TestProperty",
	//	AdditionalData = new Dictionary<string, object> { { "MyProp", "MyValue" } }
	//});

	var todoTask = new TodoTask
	{
		Title = "A new task for C#",
		LinkedResources = new TodoTaskLinkedResourcesCollectionPage()
				{
					new LinkedResource
					{
						WebUrl = "http://microsoft.com",
						ApplicationName = "Microsoft",
						DisplayName = "Microsoft"
					}
				}
	};


	var r = await client.Users[Chambele.Beta.User1].Todo.Lists[Chambele.Beta.User1TaskList].Tasks.Request().AddAsync(todoTask);

//	int i = 9 / 2;
//	Console.WriteLine(i.ToString());
//
//	int[] n = new int[] { 2, 6, 8, 10 };
//
//	var dict = new Dictionary<string, int>();
//
//	var a = "The story of the hummingbird is about this huge forest being consumed by a fire. All the animals in the forest come out and they are transfixed as they watch the forest burning and they feel very overwhelmed very powerless except this little hummingbird. It says I am going to do something about the fire. So it flies to the nearest stream and takes a drop of water. It puts it on the fire and goes up and down up and down up and down as fast as it can."
//	string[] words = a.Split(' ', a.Length);
//	foreach (string s in words)
//	{
//		if (s.IndexOf('.') > 0)
//		{
//
//		}
//	}
//var asdf = new Stack();
//asdf.pee
//
//	var b = "{7, 75, 4, 30, 1, 0}";
//	
//	await client.Connections[""].Items[""].Request().CreateAsync(new ExternalItem());
	
	/*
	var oUser = new Microsoft.Graph.User()
	{
		AccountEnabled = true, //True by default
		DisplayName = "Dummy",
		GivenName = "First name",
		Surname = "Last name",
		UserPrincipalName = "dummy@chambele.onmicrosoft.com",
		MailNickname = "Unknown",
		PasswordProfile = new PasswordProfile() { ForceChangePasswordNextSignIn = true, Password = "asdqw334sdc@#!1qsdAS" }

	};

	var result = await client.Chats.GetAllMessages().Request().GetAsync();
	await client.Users["id"].Chats.GetAllMessages().Request().GetAsync();
	result.Dump();
	*/
}

namespace Chambele
{
	public static class Beta
	{
		public static string User1TaskList = "AQMkADE4NmViZTM1LTE5NWYtNDM5NS05ZWNiLTAzOGMzMTJlZDZhOQAuAAADUMQS3-gD-kyIn7IKsO1w8AEA2zF8XWY3FkyBoomMVz1M5wAAAgESAAAA";
		public static string User1 = "michael@chambele.onmicrosoft.com";
		public static string MessageWithAttachment_Id = "AAMkADE4NmViZTM1LTE5NWYtNDM5NS05ZWNiLTAzOGMzMTJlZDZhOQBGAAAAAABQxBLf_AP_TIifsgqw7XDwBwDbMXxdZjcWTIGiiYxXPUznAAAAAAEMAADbMXxdZjcWTIGiiYxXPUznAAAtZdglAAA=";
		public static string AttachmentId = "AAMkADE4NmViZTM1LTE5NWYtNDM5NS05ZWNiLTAzOGMzMTJlZDZhOQBGAAAAAABQxBLf_AP_TIifsgqw7XDwBwDbMXxdZjcWTIGiiYxXPUznAAAAAAEMAADbMXxdZjcWTIGiiYxXPUznAAAtZdglAAABEgAQALmGlhZxBNRKhrRzIWv3B4I=";

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

		public static GraphServiceClient GetPublicGraphClient()
		{
			return new GraphServiceClient(CreateAuthorizationProvider());
		}

		private static IAuthenticationProvider CreateAuthorizationProvider()
		{
			var clientId = "6881477a-a153-4bf1-973e-60abfad70ad5";
			//var redirectUri = Util.GetPassword("redirectUriPublic");
			var tenantId = Util.GetPassword("chambele_tenantId");
			var authority = $"https://login.microsoftonline.com/{tenantId}/v2.0";

			//this specific scope means that application will default to what is defined in the application registration rather than using dynamic scopes
			List<string> scopes = new List<string>();
			scopes.Add("https://graph.microsoft.com/.default");

			PublicClientApplicationOptions options = new PublicClientApplicationOptions()
			{
				ClientId = clientId,
				TenantId = tenantId,
				//RedirectUri = redirectUri
			};

			var pca = PublicClientApplicationBuilder.CreateWithApplicationOptions(options)
													.Build();
			return new DeviceCodeFlowAuthorizationProvider(pca, scopes);
		}

		// Define other methods and classes here
		private class DeviceCodeFlowAuthorizationProvider : IAuthenticationProvider
		{
			private readonly IPublicClientApplication _application;
			private readonly List<string> _scopes;
			private string _authToken;
			public DeviceCodeFlowAuthorizationProvider(IPublicClientApplication application, List<string> scopes)
			{
				_application = application;
				_scopes = scopes;
			}
			public async Task AuthenticateRequestAsync(HttpRequestMessage request)
			{
				if (string.IsNullOrEmpty(_authToken))
				{
					var result = await _application.AcquireTokenWithDeviceCode(_scopes, callback =>
					{
						Console.WriteLine(callback.Message);
						return Task.FromResult(0);
					}).ExecuteAsync();
					_authToken = result.AccessToken;
				}
				request.Headers.Authorization = new AuthenticationHeaderValue("bearer", _authToken);
			}
		}

	}
}