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
using Mailbox.Client.Models;
using Mailbox.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Web;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace Mailbox.Client.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;
        private readonly ITokenAcquisition _tokenAcquisition;
        private readonly HttpClient _httpClient;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration, ITokenAcquisition tokenAcquistion)
        {
            _logger = logger;
            _configuration = configuration;
            _tokenAcquisition = tokenAcquistion;
            _httpClient = new HttpClient();
        }

        [Authorize]
        public IActionResult Index()
        {
            return View();
        }

        [Authorize]
        public async Task<IActionResult> GetAutomaticReplies()
        {
            MailboxSettings model = null;

            string accessToken = await _tokenAcquisition.GetAccessTokenForUserAsync(new List<string>() { _configuration.GetValue<string>("MailboxAPI:MailboxAPIScopes") });
            _httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            HttpResponseMessage response = await _httpClient.GetAsync($"{_configuration.GetValue<string>("MailboxAPI:MailboxAPIBaseAddress")}api/Mailbox/getautomaticreplies");
            if (response.IsSuccessStatusCode)
            {
                string automaticRepliesString = await response.Content.ReadAsStringAsync();
                model = JsonConvert.DeserializeObject<MailboxSettings>(automaticRepliesString);
            }
            else
            {
                string error = await response.Content.ReadAsStringAsync();
            }

            return View(model);
        }

        [Authorize]
        public async Task<IActionResult> UpdateAutomaticReplies(MailboxSettings model)
        {
            string accessToken = await _tokenAcquisition.GetAccessTokenForUserAsync(new List<string>() { _configuration.GetValue<string>("MailboxAPI:MailboxAPIScopes") });
            _httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            StringContent stringContent = new StringContent(JsonConvert.SerializeObject(model));
            stringContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            HttpResponseMessage response = await _httpClient.PatchAsync($"{_configuration.GetValue<string>("MailboxAPI:MailboxAPIBaseAddress")}api/Mailbox/updateautomaticreplies", stringContent);
            if (response.IsSuccessStatusCode)
            {
                string automaticRepliesString = await response.Content.ReadAsStringAsync();
                model = JsonConvert.DeserializeObject<MailboxSettings>(automaticRepliesString);
                ViewBag.Message = "Changes saved";
                ViewBag.Class = "alert-success";
            }
            else
            {
                ViewBag.Message = "Update failed";
                ViewBag.Class = "alert-danger";
            }

            return View("GetAutomaticReplies", model);
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
