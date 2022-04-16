using Genetec.Sdk;
using Genetec.Sdk.Entities;
using Genetec.Sdk.Events;
using Genetec.Sdk.Queries;
using Newtonsoft.Json.Linq;
using Serilog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CFM
{
    public partial class Service : ServiceBase
    {
        private Engine engine = new Engine();
        private Properties.Settings settings = Properties.Settings.Default;
        public Service()
        {
            InitializeComponent();
            Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.File(settings.LogPath, rollingInterval: RollingInterval.Day)
            .CreateLogger();
        }

        
        private void OnEngineLogonFailed(object sender, LogonFailedEventArgs e)
        {
            Log.Information(e.FormattedErrorMessage);
        }
        private void OnEngineLoggedOff(object sender, LoggedOffEventArgs e)
        {
            Log.Information("logged off");
        }
        private void OnEngineLoggedOn(object sender, LoggedOnEventArgs e)
        {
            Log.Information("logged on");
            var query = engine.ReportManager.CreateReportQuery(ReportType.EntityConfiguration) as EntityConfigurationQuery;
            query.EntityTypeFilter.Add(EntityType.Door);
            query.StrictResults = true;
            query.DownloadAllRelatedData = true;
            var entities = query.Query();
            List<Guid> entitiesGuids = entities.Data.Rows.Cast<DataRow>().Select(row => (Guid)row[0]).ToList();
        }
        private void OnEngineLogonStatusChanged(object sender, LogonStatusChangedEventArgs e)
        {
            Log.Information(e.Status.ToString());
        }
        private async void OnEngineEventReceived(object sender, EventReceivedEventArgs e)
        {
            switch (e.EventType)
            {
                case EventType.AccessGranted:
                    var tz = e.Event as SupportsTimeZoneEvent;
                    string badgeField = null;
                    CardholderAccessRequestedEventArgs e2 = e as CardholderAccessRequestedEventArgs;
                    DateTime time = TimeZoneInfo.ConvertTimeFromUtc(e.Timestamp, tz.SourceTimeZone);
                    AccessPoint ap = engine.GetEntity(e2.AccessPointGuid) as AccessPoint;
                    Door door = engine.GetEntity(ap.Door) as Door;
                    Cardholder employee = engine.GetEntity(e2.CardholderGuid) as Cardholder;
                    switch (settings.GT_BadgeFieldSource)
                    {
                        case "Credential":
                            badgeField = e2.GetType().GetProperty(settings.GT_BadgeField).GetValue(e2).ToString();
                            break;
                        case "CardHolder":
                            badgeField = employee.GetType().GetProperty(settings.GT_BadgeField).GetValue(employee).ToString();
                            break;
                        case "CardHolderCustomField":
                            badgeField = employee.GetCustomFields().Where(x => x.CustomField.Name == settings.GT_BadgeField).FirstOrDefault().Value as string;
                            break;

                    }
                    await CardAdmitted(badgeField, time.ToString(), door.Name);
                    //Task.Run(async () => await CardAdmitted(e2.CardholderGuid.ToString(), time.ToString(), door.Name));
                    break;
            }
        }
        protected override void OnStart(string[] args)
        {
            SubscribeEngine();
            if (settings.GT_AuthType == "Windows")
                engine.LogOnUsingWindowsCredential(settings.GT_server);
            if (settings.GT_AuthType == "Digest")
                engine.LogOn(settings.GT_server, settings.GT_username, settings.GT_password);
        }
        protected override void OnStop()
        {
            engine.LogOff();
            Dispose();
        }

        private new void Dispose()
        {
            UnsubscribeEngine();
            engine.Dispose();
            engine = null;
            base.Dispose();
        }
        private void SubscribeEngine()
        {
            engine.LoggedOff += OnEngineLoggedOff;
            engine.LoggedOn += OnEngineLoggedOn;
            engine.LogonFailed += OnEngineLogonFailed;
            engine.LogonStatusChanged += OnEngineLogonStatusChanged;
            engine.EventReceived += OnEngineEventReceived;
            
        }
        private void UnsubscribeEngine()
        {
            engine.LoggedOn -= OnEngineLoggedOn;
            engine.LoggedOff -= OnEngineLoggedOff;
            engine.LogonFailed -= OnEngineLogonFailed;
            engine.LogonStatusChanged -= OnEngineLogonStatusChanged;
            engine.EventReceived -= OnEngineEventReceived;
        }

        public async Task CardAdmitted(string cardid, string timestamp, string door)
        {
            try
            {
                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {settings.CFM_Token}");
                var values = new Dictionary<string, string>{
                    { "class", "cfm.wr.BadgingService" },
                    { "method", "CardAdmitted" },
                    { "st_cardid", cardid },
                    { "st_timestamp", timestamp },
                    { "st_door", door }
                };
                var response = await client.PostAsync(settings.CFM_EndPoint, new FormUrlEncodedContent(values));
                if (false == response.IsSuccessStatusCode)
                {
                    Log.Information(await response.Content.ReadAsStringAsync());
                }
                JObject res = JObject.Parse(await response.Content.ReadAsStringAsync());
                Log.Information("CadapultFM response: " + (string)res["msg"]);
            }
            catch (HttpRequestException ex)
            {
                Log.Information("Request Error: "+ex.Message);
            }
        }
    }
}
