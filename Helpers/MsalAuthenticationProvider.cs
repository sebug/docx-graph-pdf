using System.Net.Http.Headers;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace docx_graph_pdf.Helpers;

public class MsalAuthenticationProvider : IAuthenticationProvider
{
    private readonly IConfidentialClientApplication _clientApplication;
    private readonly string[] _scopes;

    public MsalAuthenticationProvider(IConfidentialClientApplication clientApplication,
    string[] scopes)
    {
        _clientApplication = clientApplication;
        _scopes = scopes;
    }
    

    public async Task AuthenticateRequestAsync(HttpRequestMessage request)
    {
        var token = await GetTokenAsync();
        request.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
    }

    public async Task<string> GetTokenAsync()
    {
        AuthenticationResult authResult;
        authResult = await _clientApplication.AcquireTokenForClient(_scopes).ExecuteAsync();
        return authResult.AccessToken;
    }
}