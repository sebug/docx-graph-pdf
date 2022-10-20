// See https://aka.ms/new-console-template for more information
using docx_graph_pdf;
using Microsoft.Extensions.Configuration;

var config = new ConfigurationBuilder()
    .AddEnvironmentVariables()
    .Build();

var settings = config.GetRequiredSection("DocxGraphPdf").Get<DocxGraphPdfOptions>();

Console.WriteLine("Hello, " + settings.RedirectUri);
