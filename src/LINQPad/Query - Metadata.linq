<Query Kind="Statements">
  <Namespace>System.Net</Namespace>
</Query>

string v1_0 = "https://graph.microsoft.com/v1.0/$metadata";
string beta = "https://graph.microsoft.com/beta/$metadata";
string cleaned = "https://raw.githubusercontent.com/microsoftgraph/msgraph-metadata/master/clean_v10_metadata/cleanMetadataWithDescriptionsv1.0.xml";
string metadata = new WebClient().DownloadString(beta);
XElement xMetadata = XElement.Parse(metadata);

// Get all of the used Org.OData.Capabilities.V1 terms in the metadata.
//var capaTerms = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Annotation").Attributes("Term").Where(t => t.Value.Contains("Org.OData.Capabilities.V1"));
//capaTerms.Dump();

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

// Get all elements
//var result = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Schema").Elements();
//result.Dump();

// Get all entities that have abstract base type referenced in the navigation of another entity.
//var results = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}EntityType").Where(m => m.Attributes("Abstract").Any());
//results.Dump("Entities with base type");

//var results = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}ComplexType").Where(m => m.Attributes("Abstract").Any());
//results.Dump("Abstract complex types");

// Get all actions
//var allActions = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Action");
//allActions.Dump(nameof(allActions));

// Get all actions bound to a type 
/*
xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Action")
		 .Where(m => m.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Parameter")
		 	.First()
			.Attributes("Type")
			.Any(x => x.Value.Equals("graph.workbookRange")))
		 .Dump("All actions to bound workbookRange");
*/

// Get all functions bound to a type
/*
xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Function")
		 .Where(m => m.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Parameter")
		 	.First()
			.Attributes("Type")
			.Any(x => x.Value.Equals("graph.workbookRange")))
		 .Dump("All functions to bound workbookRange");


xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Function")
		 .Where(m => m.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Parameter")
		 	.First()
			.Attributes("Type")
			.Any(x => x.Value.Contains("graph.workbookChart")))
		 .Dump("All functions to bound workbookChart");
*/		 

// Get all functions
//var allFunctions = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Function");
//allFunctions.Dump(nameof(allFunctions));

/*
// Get all actions with a OData primitive return types (not stream)
var results = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Action").Where(m => m.Descendants("{http://docs.oasis-open.org/odata/ns/edm}ReturnType")
																									 .Attributes("Type")
																									 .Any(x => x.Value.StartsWith("Edm.") && !x.Value.StartsWith("Edm.Stream")));

var results2 = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Function").Where(m => m.Descendants("{http://docs.oasis-open.org/odata/ns/edm}ReturnType")
																									 .Attributes("Type")
																									 .Any(x => x.Value.StartsWith("Edm.") && !x.Value.StartsWith("Edm.Stream")));

var results3 = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Action").Where(m => m.Descendants("{http://docs.oasis-open.org/odata/ns/edm}ReturnType")
																									 .Attributes("Type")
																									 .Any(x => x.Value.StartsWith("Collection(Edm.")&& !x.Value.StartsWith("Collection(Edm.Stream")));

var results4 = xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Function").Where(m => m.Descendants("{http://docs.oasis-open.org/odata/ns/edm}ReturnType")
																									 .Attributes("Type")
																									 .Any(x => x.Value.StartsWith("Collection(Edm.")&& !x.Value.StartsWith("Collection(Edm.Stream")));

results.Dump("Action with primitive");
results2.Dump("Function with primitive");
results3.Dump("Action with primitive collection");
results4.Dump("Functions with primitive collection");
*/

// Find all composable functions
//xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Function")
//		 .Where(m => m.Attributes("IsComposable").Any())
//         .Dump("All composable functions");

xMetadata.Descendants("{http://docs.oasis-open.org/odata/ns/edm}Function")
		 .Where(m => m.Attributes("IsComposable").Any())
		 .Where(m => m.Descendants("{http://docs.oasis-open.org/odata/ns/edm}ReturnType")
			.Attributes("Type")
			.Any(x => !x.Value.Contains("graph.report")))
		 .Dump("All composable functions without report root");
		 

