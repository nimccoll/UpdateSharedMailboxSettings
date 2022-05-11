//===============================================================================
// Microsoft FastTrack for Azure
// Update Shared Mailbox Automatic Replies Sample
//===============================================================================
// Copyright © Microsoft Corporation.  All rights reserved.
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY
// OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
// LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
// FITNESS FOR A PARTICULAR PURPOSE.
//===============================================================================
using Mailbox.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace Mailbox.API.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class MailboxController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private readonly HttpClient _httpClient;

        public MailboxController(IConfiguration configuration)
        {
            _configuration = configuration;
            _httpClient = new HttpClient();
        }

        // GET: api/mailbox/getautomaticreplies
        [Route("getautomaticreplies")]
        [HttpGet]
        public async Task<IActionResult> Get()
        {
            HttpResponseMessage automaticRepliesResponse = new HttpResponseMessage();

            // Retrieve access token for the Microsoft Graph using the Client Credential flow
            IConfidentialClientApplication app;
            app = ConfidentialClientApplicationBuilder.Create(_configuration.GetValue<string>("AzureAd:ClientId"))
                                                      .WithClientSecret(_configuration.GetValue<string>("AzureAd:ClientSecret"))
                                                      .WithAuthority(new Uri($"{_configuration.GetValue<string>("AzureAd:Instance")}{_configuration.GetValue<string>("AzureAd:TenantId")}"))
                                                      .Build();
            string[] scopes = new string[] { $"{_configuration.GetValue<string>("GraphEndpoint")}{_configuration.GetValue<string>("GraphScope")}" };
            AuthenticationResult authenticationResult = await app.AcquireTokenForClient(scopes).ExecuteAsync();

            // Retrieve the Automatic Replies settings for the share mailbox
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authenticationResult.AccessToken);
            HttpResponseMessage httpResponseMessage = await _httpClient.GetAsync($"{_configuration.GetValue<string>("GraphEndpoint")}v1.0/users/{_configuration.GetValue<string>("Mailbox")}/mailboxSettings");
            if (httpResponseMessage.IsSuccessStatusCode)
            {
                automaticRepliesResponse.StatusCode = System.Net.HttpStatusCode.OK;
                string automaticRepliesResponseString = await httpResponseMessage.Content.ReadAsStringAsync();
                return new OkObjectResult(automaticRepliesResponseString);
            }
            else
            {
                return new StatusCodeResult((int)httpResponseMessage.StatusCode);
            }
        }

        // GET: api/mailbox/updateautomaticreplies
        [Route("updateautomaticreplies")]
        [HttpPatch]
        public async Task<IActionResult> Patch()
        {
            HttpResponseMessage automaticRepliesResponse = new HttpResponseMessage();

            // Verify that the body contains a MailboxSettings object
            StreamReader streamReader = new StreamReader(this.Request.Body);
            string updatedMailboxSettingsString = await streamReader.ReadToEndAsync();
            MailboxSettings updatedMailboxSettings = JsonConvert.DeserializeObject<MailboxSettings>(updatedMailboxSettingsString);
            streamReader.Close();

            // Retrieve access token for the Microsoft Graph using the Client Credential flow
            IConfidentialClientApplication app;
            app = ConfidentialClientApplicationBuilder.Create(_configuration.GetValue<string>("AzureAd:ClientId"))
                                                      .WithClientSecret(_configuration.GetValue<string>("AzureAd:ClientSecret"))
                                                      .WithAuthority(new Uri($"{_configuration.GetValue<string>("AzureAd:Instance")}{_configuration.GetValue<string>("AzureAd:TenantId")}"))
                                                      .Build();
            string[] scopes = new string[] { $"{_configuration.GetValue<string>("GraphEndpoint")}{_configuration.GetValue<string>("GraphScope")}" };
            AuthenticationResult authenticationResult = await app.AcquireTokenForClient(scopes).ExecuteAsync();

            // Update the Automatic Replies settings on the shared mailbox
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authenticationResult.AccessToken);
            StringContent stringContent = new StringContent(JsonConvert.SerializeObject(updatedMailboxSettings));
            stringContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            HttpResponseMessage httpResponseMessage = await _httpClient.PatchAsync($"{_configuration.GetValue<string>("GraphEndpoint")}v1.0/users/{_configuration.GetValue<string>("Mailbox")}/mailboxSettings", stringContent);
            if (httpResponseMessage.IsSuccessStatusCode)
            {
                automaticRepliesResponse.StatusCode = System.Net.HttpStatusCode.OK;
                string automaticRepliesResponseString = await httpResponseMessage.Content.ReadAsStringAsync();
                return new OkObjectResult(automaticRepliesResponseString);
            }
            else
            {
                return new StatusCodeResult((int)httpResponseMessage.StatusCode);
            }
        }
    }
}
