// See https://aka.ms/new-console-template for more information
using docx_graph_pdf;
using docx_graph_pdf.Helpers;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Identity.Client;

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

Console.WriteLine("folder created: " + r.WebUrl);

GraphServiceClient GetAuthenticatedGraphClient(DocxGraphPdfOptions options)
{
    var authenticationProvider = CreateAuthorizationProvider(options);
    var graphClient = new GraphServiceClient(authenticationProvider);
    return graphClient;
}

IAuthenticationProvider CreateAuthorizationProvider(DocxGraphPdfOptions options)
{
    var scopes = new List<string>()
    {
        "https://graph.microsoft.com/.default"
    };
    var authority = $"https://login.microsoftonline.com/{options.TenantID}/v2.0";
    var cca = ConfidentialClientApplicationBuilder.Create(options.ApplicationID)
        .WithAuthority(authority)
        .WithRedirectUri(options.RedirectUri)
        .WithClientSecret(options.ApplicationSecret)
        .Build();
    
    return new MsalAuthenticationProvider(cca, scopes.ToArray());
}