<Query Kind="Statements">
  <Namespace>System.Net</Namespace>
</Query>

string v1_0 = "https://graph.microsoft.com/v1.0/$metadata";
string beta = "https://graph.microsoft.com/beta/$metadata";
string cleaned = "https://raw.githubusercontent.com/microsoftgraph/msgraph-metadata/master/clean_v10_metadata/cleanMetadataWithDescriptionsv1.0.xml";
string metadata = new WebClient().DownloadString(v1_0);
XElement xMetadata = XElement.Parse(metadata);

/* Works

// Get all annotations including target.
var results = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Annotations");

// Get all annotation elements.
var results = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Annotation");

// Get all of the used terms in the metadata.
var results = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Annotation").Attributes("Term");

// Get all ofthe distinct terms in the metadata.
var results = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Annotation").Attributes("Term")
																						  .Select(m => m.Value)
																						  .Distinct();

// Get a count of each distinct term
var results = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Annotation").Attributes("Term")
																						  .GroupBy(m => m.Value)
																						  .Select(n => new {
																						  	Term = n.Select(t => t.Value),
																							Count = n.Select(m => m.Value).Distinct().Count()
																						  });
																						  
// Get the count of complex types
var results = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}ComplexType").ToList().Count;

// Get all complex types with a defined base type. 39
var results = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}ComplexType").Where(m => m.Attributes("BaseType").Any());

// Get all complex types. 283
var results = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}ComplexType");

// Get all abstract complex types. 9
var results = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}ComplexType").Where(m => m.Attributes("Abstract").Any());

*/

// Get all entity types. 319
//var results = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}EntityType");

// Get all abstract entity types. 28
// var results = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}EntityType").Where(m => m.Attributes("Abstract").Any());

// Get all actions with a OData primitive return types (not stream)
var results = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Action").Where(m => m.Descendants("{http://docs.oasis-open.org/odata/ns/edm}ReturnType")
																									 .Attributes("Type")
																									 .Any(x => x.Value.StartsWith("Edm.") && !x.Value.StartsWith("Edm.Stream")));

var results2 = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Function").Where(m => m.Descendants("{http://docs.oasis-open.org/odata/ns/edm}ReturnType")
																									 .Attributes("Type")
																									 .Any(x => x.Value.StartsWith("Edm.") && !x.Value.StartsWith("Edm.Stream")));



results.Dump();
results2.Dump();