<Query Kind="Program">
  <NuGetReference Prerelease="true">Microsoft.Graph.Beta</NuGetReference>
  <Namespace>Microsoft.Graph</Namespace>
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
  <Namespace>System.Runtime.Serialization.Formatters</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>

async Task Main()
{
	var client = Beta.GetConfidentialClient();

	await Temp(client);
}

static async Task Temp(Microsoft.Graph.GraphServiceClient client)
{

}
