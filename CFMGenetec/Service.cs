using Genetec.Sdk;
using Genetec.Sdk.Entities;
using Genetec.Sdk.Events;
using Genetec.Sdk.Queries;
using Newtonsoft.Json;
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
        private System.Timers.Timer doortimer = new System.Timers.Timer();
        private Dictionary<string, string> doors = new Dictionary<string, string>();
        public Service()
        {
            InitializeComponent();
            Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.File(settings.LogPath, rollingInterval: RollingInterval.Day)
            .CreateLogger();
            doortimer.Elapsed += async (e, a) => { await refreshDoorsData(); };
            doortimer.Interval = 7200000;//2 hours
            doortimer.Enabled = true;
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
            try
            {
                switch (e.EventType)
                {
                    case EventType.AccessGranted:
                        var tz = e.Event as SupportsTimeZoneEvent;
                        string badgeField = null;
                        CardholderAccessRequestedEventArgs e2 = e as CardholderAccessRequestedEventArgs;
                        DateTime time = TimeZoneInfo.ConvertTimeFromUtc(e.Timestamp, tz.SourceTimeZone);
                        AccessPoint ap = engine.GetEntity(e2.AccessPointGuid) as AccessPoint;
                        var door = engine.GetEntity(ap.Door) as Genetec.Sdk.Entities.Door;
                        Cardholder employee = engine.GetEntity(e2.CardholderGuid) as Cardholder;
                        if (null == employee) return;
                        switch (settings.GT_BadgeFieldSource)
                        {
                            case "Credential":
                                Log.Information("Event recieved: Credential");
                                badgeField = e2.GetType().GetProperty(settings.GT_BadgeField).GetValue(e2).ToString();
                                break;
                            case "CardHolder":
                                Log.Information("Event recieved: CardHolder");
                                badgeField = employee.GetType().GetProperty(settings.GT_BadgeField).GetValue(employee).ToString();
                                break;
                            case "CardHolderCustomField":
                                Log.Information("Event recieved: CardHolderCustomField");
                                badgeField = employee.GetCustomFields().Where(x => x.CustomField.Name == settings.GT_BadgeField).FirstOrDefault().Value as string;
                                break;
                            default:
                                return;
                        }
                        Log.Information($"Calling CardAdmitted {badgeField},{time.ToString()}, {door.Name}");
                        await CardAdmitted(badgeField, time.ToString(), door.Name);
                        //Task.Run(async () => await CardAdmitted(e2.CardholderGuid.ToString(), time.ToString(), door.Name));
                        break;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
            }

        }
        protected override async void OnStart(string[] args)
        {
            Log.Information("On Start called");
            await refreshDoorsData();
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
                if (!doors.ContainsKey(door.ToLower()))
                {
                    Log.Information($"CardAdmitted: {door} not found");
                    return;
                }
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
            catch (Exception ex)
            {
                Log.Error(ex,"");
            }
        }
        public async Task refreshDoorsData()
        {
            try
            {
                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {settings.CFM_Token}");
                var values = new Dictionary<string, string>{
                    { "class", "cfm.wr.AdminService" },
                    { "method", "getAllBadgeDoorsWithLocation" }
                };
                var response = await client.PostAsync(settings.CFM_AdminEndPoint, new FormUrlEncodedContent(values));
                
                if (false == response.IsSuccessStatusCode)
                {
                    Log.Error(await response.Content.ReadAsStringAsync());
                }
                var res = await response.Content.ReadAsStringAsync();
                doors = JsonConvert.DeserializeObject<List<Door>>(res).Aggregate(new Dictionary<string, string>(), (t, n) =>
                {
                    n.doorid = n.doorid.ToLower();
                    if(!t.ContainsKey(n.doorid))
                        t.Add(n.doorid, n.timezone);
                    return t;
                });
                Log.Information("Door data refreshed");
            }
            catch (Exception ex)
            {
                Log.Error(ex,"");
            }
        }
    }
    class Door
    {
        public string doorid;
        public string timezone;
    }
}
