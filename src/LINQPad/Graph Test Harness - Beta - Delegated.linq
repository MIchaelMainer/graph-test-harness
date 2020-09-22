<Query Kind="Program">
  <NuGetReference Prerelease="true">Microsoft.Graph.Beta</NuGetReference>
  <NuGetReference>Microsoft.Identity.Client</NuGetReference>
  <Namespace>Microsoft.Graph</Namespace>
  <Namespace>Microsoft.Identity.Client</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Net.Http.Headers</Namespace>
  <Namespace>System.Text.Json</Namespace>
  <Namespace>System.Text.Json.Serialization</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>

private static GraphServiceClient _graphClient;

async void Main()
{
	GraphServiceClient graphClient = GetAuthenticatedGraphClient();
	
	//graphClient.Users.Request().GetAsync().Result[0].DisplayName.Dump("Graph result");


	var externalConnection = new ExternalConnection
	{
		Id = "contosohr",
		Name = "Contoso HR",
		Description = "Description"
	};

    await graphClient.Reports.GetOneDriveActivityUserDetail("D7").Request().GetAsync()

string result = await graphClient.Education
								 .SynchronizationProfiles[""]
								 .UploadUrl()
								 .Request()
								 .GetAsync();
	
IDirectoryObjectCheckMemberGroupsCollectionPage page = await graphClient.Me
																		.CheckMemberGroups(new List<string>())
																		.Request()
																		.PostAsync();

}

private GraphServiceClient GetAuthenticatedGraphClient()
{
	var authenticationProvider = CreateAuthorizationProvider();
	_graphClient = new GraphServiceClient(authenticationProvider);
	return _graphClient;
}

private static IAuthenticationProvider CreateAuthorizationProvider()
{
	var clientId = Util.GetPassword("clientIdPublic");
	var redirectUri = Util.GetPassword("redirectUriPublic");
	var tenantId = Util.GetPassword("tenantId");
	var authority = $"https://login.microsoftonline.com/{tenantId}/v2.0";

	//this specific scope means that application will default to what is defined in the application registration rather than using dynamic scopes
	List<string> scopes = new List<string>();
	scopes.Add("https://graph.microsoft.com/.default");

	PublicClientApplicationOptions options = new PublicClientApplicationOptions()
	{
		ClientId = clientId,
		TenantId = tenantId,
		RedirectUri = redirectUri,
	};

	var pca = PublicClientApplicationBuilder.CreateWithApplicationOptions(options)
											.Build();
	return new DeviceCodeFlowAuthorizationProvider(pca, scopes);
}

// Define other methods and classes here
public class DeviceCodeFlowAuthorizationProvider : IAuthenticationProvider
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

public class AuthHandler : DelegatingHandler
{
	private IAuthenticationProvider _authenticationProvider;

	public AuthHandler(IAuthenticationProvider authenticationProvider, HttpMessageHandler innerHandler)
	{
		InnerHandler = innerHandler;
		_authenticationProvider = authenticationProvider;
	}

	protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
	{
		await _authenticationProvider.AuthenticateRequestAsync(request);
		return await base.SendAsync(request, cancellationToken);
	}
}