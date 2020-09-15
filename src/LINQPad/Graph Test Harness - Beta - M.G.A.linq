<Query Kind="Program">
  <NuGetReference Prerelease="true">Microsoft.Graph.Beta</NuGetReference>
  <NuGetReference Prerelease="true">Microsoft.Graph.Auth</NuGetReference>
  <NuGetReference>Microsoft.Identity.Client</NuGetReference>
  <Namespace>Microsoft.Graph</Namespace>
  <Namespace>Microsoft.Graph.Auth</Namespace>
  <Namespace>Microsoft.Identity.Client</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>


/// <summary>Get beta generated clients using Microsoft.Graph.Auth.</summary>
public static class Beta
{
	public static Microsoft.Graph.GraphServiceClient GetConfidentialClient()
	{
		var authClient = ConfidentialClientApplicationBuilder
						   .Create(Util.GetPassword("clientId"))
						   .WithTenantId(Util.GetPassword("tenantId"))
						   .WithClientSecret(Util.GetPassword("clientsecret"))
						   .Build();

		ClientCredentialProvider authenticationProvider = new ClientCredentialProvider(authClient);
		return new GraphServiceClient(authenticationProvider);
	}

	/// <summary>Gets a public client application. You will need to use the Request() .WithUsernamePassword(email, password)</summary>
	public static Microsoft.Graph.GraphServiceClient GetPublicClient()
	{
		var authClient = PublicClientApplicationBuilder
							.Create(Util.GetPassword("username_password_clientId"))
							.Build();

		UsernamePasswordProvider authenticationProvider = new UsernamePasswordProvider(authClient);

		return new GraphServiceClient(authenticationProvider);
	}
}
