// See https://aka.ms/new-console-template for more information
using System.Net.Http.Headers;
using docx_graph_pdf;
using docx_graph_pdf.Helpers;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Identity.Client;

if (args.Length == 0)
{
    throw new Exception("Usage: docx-graph-pdf wordFile.docx");
}

var config = new ConfigurationBuilder()
    .AddEnvironmentVariables()
    .Build();

var settings = config.GetRequiredSection("DocxGraphPdf").Get<DocxGraphPdfOptions>();

var client = GetAuthenticatedGraphClient(settings);

var folderId = "Convert_" + Guid.NewGuid().ToString().Replace("-", String.Empty);

var item = new DriveItem
{
    Name = folderId,
    Folder = new Folder(),
    AdditionalData = new Dictionary<string, object>()
    {
        {"@microsoft.graph.conflictBehavior","rename"}
    }
};

var r = await client.Drive.Root.Children.Request().AddAsync(item);

var uploadSession = await client.Drive.Root.ItemWithPath(r.Name + "/" + System.IO.Path.GetFileName(args[0]))
                        .CreateUploadSession().Request().PostAsync();

var largeFileUploadTask = new LargeFileUploadTask<DriveItem>(uploadSession, System.IO.File.OpenRead(args[0]));

var uploadResponse = await largeFileUploadTask.UploadAsync();

var requestUrl = client.Drive.Root.ItemWithPath(r.Name + "/" + System.IO.Path.GetFileName(args[0])).RequestUrl + "/content?format=pdf";

var cca = BuildConfidentialClientApplication(settings);

var tokenResponse = await cca.AcquireTokenForClient(new List<string>()
{
    "https://graph.microsoft.com/.default"
}).ExecuteAsync();

var httpClient = new HttpClient();
httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", tokenResponse.AccessToken);

var convertResponse = await httpClient.GetAsync(requestUrl);

var responseBytes = await convertResponse.Content.ReadAsByteArrayAsync();

System.IO.File.WriteAllBytes(args[0].Replace(".docx", ".pdf"), responseBytes);

GraphServiceClient GetAuthenticatedGraphClient(DocxGraphPdfOptions options)
{
    var authenticationProvider = CreateAuthorizationProvider(options);
    var graphClient = new GraphServiceClient(authenticationProvider);
    return graphClient;
}

IConfidentialClientApplication BuildConfidentialClientApplication(DocxGraphPdfOptions options)
{
    var authority = $"https://login.microsoftonline.com/{options.TenantID}/v2.0";
    var cca = ConfidentialClientApplicationBuilder.Create(options.ApplicationID)
        .WithAuthority(authority)
        .WithRedirectUri(options.RedirectUri)
        .WithClientSecret(options.ApplicationSecret)
        .Build();
    return cca;
}


IAuthenticationProvider CreateAuthorizationProvider(DocxGraphPdfOptions options)
{
    var scopes = new List<string>()
    {
        "https://graph.microsoft.com/.default"
    };
    var cca = BuildConfidentialClientApplication(options);
    
    return new MsalAuthenticationProvider(cca, scopes.ToArray());
}