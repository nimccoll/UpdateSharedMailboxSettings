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
using Newtonsoft.Json;
using System;

namespace Mailbox.Models
{
    // MailboxSettings myDeserializedClass = JsonConvert.DeserializeObject<MailboxSettings>(myJsonResponse);
    public class AutomaticRepliesSetting
    {
        public string status { get; set; }
        public string externalAudience { get; set; }
        public string internalReplyMessage { get; set; }
        public string externalReplyMessage { get; set; }
        public ScheduledStartDateTime scheduledStartDateTime { get; set; }
        public ScheduledEndDateTime scheduledEndDateTime { get; set; }
    }

    public class MailboxSettings
    {
        [JsonProperty("@odata.context")]
        public string OdataContext { get; set; }
        public AutomaticRepliesSetting automaticRepliesSetting { get; set; }
    }

    public class ScheduledEndDateTime
    {
        public DateTime dateTime { get; set; }
        public string timeZone { get; set; }
    }

    public class ScheduledStartDateTime
    {
        public DateTime dateTime { get; set; }
        public string timeZone { get; set; }
    }

    public class TimeZone
    {
        public string name { get; set; }
    }
}
