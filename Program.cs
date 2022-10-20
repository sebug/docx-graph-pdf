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

Console.WriteLine("Hello, " + settings.RedirectUri);

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