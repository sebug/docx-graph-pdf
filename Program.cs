﻿// See https://aka.ms/new-console-template for more information
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

Console.WriteLine("file created: " + requestUrl);

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