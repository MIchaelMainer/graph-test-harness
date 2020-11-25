<Query Kind="Program">
  <IncludeUncapsulator>false</IncludeUncapsulator>
</Query>


/*
void Main()
{
	// Write code to test your extensions here. Press F5 to compile and run.
}

public static class MyExtensions
{
	// Write custom extension methods here. They will be available to all queries.
	
}

// Get clients setup to use the chambele tenant. Only uses the v1 client.
namespace Chambele
{
	public static class V1
	{
		public static Microsoft.Graph.GraphServiceClient GetConfidentialClient()
		{
			var authClient = ConfidentialClientApplicationBuilder
							   .Create(Util.GetPassword("chambele_clientId"))
							   .WithTenantId(Util.GetPassword("chambele_tenantId"))
							   .WithClientSecret(Util.GetPassword("chambele_clientsecret"))
							   .Build();

			ClientCredentialProvider authenticationProvider = new ClientCredentialProvider(authClient);
			return new GraphServiceClient(authenticationProvider);
		}
	}
}

// Get clients setup to use the M365x462896 tenant. Only uses the v1 client.
namespace M365x462896
{
	/// <summary>Get V1 generated clients using Microsoft.Graph.Auth.</summary>
	public static class V1
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
}
*/


#region Advanced - How to multi-target

// The NET5 symbol can be useful when you want to run some queries under .NET 5 and others under .NET Core 3:

#if NET5
// Code that requires .NET 5 or later
#endif

#endregion