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
</Query>

async Task Main()
{
	var client = Chambele.Beta.GetConfidentialGraphClient();

	await Temp(client);
}

static async Task Temp(Microsoft.Graph.GraphServiceClient client)
{
	var ser = new Serializer();
	var bodyContent = ser.SerializeObject(new ExternalItem());
	
	var externalItemRequest = client.Connections[""].Items[""].Request();
 	var httpRequestMessage = externalItemRequest.GetHttpRequestMessage();
	httpRequestMessage.Method = HttpMethod.Put;
	httpRequestMessage.Content = new StringContent(bodyContent);
	 
	var httpResponseMessage = await externalItemRequest.Client.HttpProvider.SendAsync(httpRequestMessage);
	var externItem = ser.DeserializeObject<ExternalItem>(httpResponseMessage.Content.ReadAsStringAsync().Result);

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
		public static string user = "michael@chambele.onmicrosoft.com";
		public static string messageWithAttachment_Id = "AAMkADE4NmViZTM1LTE5NWYtNDM5NS05ZWNiLTAzOGMzMTJlZDZhOQBGAAAAAABQxBLf_AP_TIifsgqw7XDwBwDbMXxdZjcWTIGiiYxXPUznAAAAAAEMAADbMXxdZjcWTIGiiYxXPUznAAAtZdglAAA=";
		public static string attachmentId = "AAMkADE4NmViZTM1LTE5NWYtNDM5NS05ZWNiLTAzOGMzMTJlZDZhOQBGAAAAAABQxBLf_AP_TIifsgqw7XDwBwDbMXxdZjcWTIGiiYxXPUznAAAAAAEMAADbMXxdZjcWTIGiiYxXPUznAAAtZdglAAABEgAQALmGlhZxBNRKhrRzIWv3B4I=";

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