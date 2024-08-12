using Newtonsoft.Json;
using System.Configuration;
using System.Diagnostics.Eventing.Reader;
using System.Text;
using System.Windows.Forms;

namespace SamWafWinEvent
{
    internal static class Program
    {
        private static string AppId = "";
        private static string AppSecret = "";
        private static string SendToOpenId = "";
        private static string SendToTemplateId = "";
        private static string CheckProviderName = "";
        private static readonly string ConfigFilePath = "config.ini";

        private static readonly string TokenFilePath = "access_token.json";
        private static readonly string LogFilePath = "notification_log.json";
        //每天通知的最大值
        private static readonly int MaxNotificationsPerDay = 10;
        private static readonly TimeSpan NotificationInterval = TimeSpan.FromMinutes(1);
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (args is null) throw new ArgumentNullException(nameof(args));
            LoadConfig();
            Console.WriteLine($"【配置信息读取情况】:");
            Console.WriteLine($"AppId: {MaskValue(AppId)}");
            Console.WriteLine($"AppSecret: {MaskValue(AppSecret)}");
            Console.WriteLine($"SendToOpenId: {MaskValue(SendToOpenId)}");
            Console.WriteLine($"SendToTemplateId: {MaskValue(SendToTemplateId)}");

            LoadEventLogs(); 
            Console.ReadKey();
        }
        private static string MaskValue(string value)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= 4)
            {
                return new string('*', value.Length);
            }

            int visibleCount = Math.Min(4, value.Length / 4);
            string maskedPart = new string('*', value.Length - visibleCount);
            string visiblePart = value.Substring(value.Length - visibleCount);
            return maskedPart + visiblePart;
        }
        private static void LoadConfig()
        {
            if (!File.Exists(ConfigFilePath))
            {
                CreateDefaultConfig();
            }

            var lines = File.ReadAllLines(ConfigFilePath);
            foreach (var line in lines)
            {
                if (line.StartsWith("AppId ="))
                {
                    AppId = line.Substring("AppId =".Length).Trim();
                }
                else if (line.StartsWith("AppSecret ="))
                {
                    AppSecret = line.Substring("AppSecret =".Length).Trim();
                }
                else if (line.StartsWith("SendToOpenId ="))
                {
                    SendToOpenId = line.Substring("SendToOpenId =".Length).Trim();
                }
                else if (line.StartsWith("SendToTemplateId ="))
                {
                    SendToTemplateId = line.Substring("SendToTemplateId =".Length).Trim();
                }
            }
        }

        private static void CreateDefaultConfig()
        {
            using (StreamWriter writer = new StreamWriter(ConfigFilePath))
            {
                writer.WriteLine("[WeChat]");
                writer.WriteLine("AppId = ");
                writer.WriteLine("AppSecret = ");
                writer.WriteLine("SendToOpenId = ");
                writer.WriteLine("SendToTemplateId = ");
            }
        }
        static async Task<string> GetAccessTokenAsync()
        {
            if (File.Exists(TokenFilePath))
            {
                var tokenData = File.ReadAllText(TokenFilePath);
                var token = JsonConvert.DeserializeObject<TokenResponse>(tokenData);

                if (token != null && token.Expiry > DateTime.UtcNow)
                {
                    return token.AccessToken;
                }
            }

            var newToken = await FetchAccessTokenAsync();
            File.WriteAllText(TokenFilePath, JsonConvert.SerializeObject(newToken));
            return newToken.AccessToken;
        }

        private static async Task<TokenResponse> FetchAccessTokenAsync()
        {
            var url = $"https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid={AppId}&secret={AppSecret}";

            using (var client = new HttpClient())
            {
                var response = await client.GetStringAsync(url);
                var token = JsonConvert.DeserializeObject<TokenResponse>(response);

                if (token != null)
                {
                    token.Expiry = DateTime.UtcNow.AddSeconds(token.ExpiresIn - 100); // 缓存时间略短于实际过期时间
                }

                return token;
            }
        }

        private class TokenResponse
        {
            [JsonProperty("access_token")]
            public string AccessToken { get; set; }

            [JsonProperty("expires_in")]
            public int ExpiresIn { get; set; }

            [JsonIgnore]
            public DateTime Expiry { get; set; }
        }
        private class NotificationLog
        {
            public DateTime Timestamp { get; set; }
        }

        private class NotificationMessage
        {
            [JsonProperty("touser")]
            public string ToUser { get; set; }

            [JsonProperty("template_id")]
            public string TemplateId { get; set; }

            [JsonProperty("data")]
            public object Data { get; set; }
        }

        private static void LoadEventLogs()
        {
            EventLogSession session = new EventLogSession(); 
            EventLogQuery query = new EventLogQuery("Application", PathType.LogName)
            {
                TolerateQueryErrors = true,
                Session = session
            };

            EventLogWatcher logWatcher = new EventLogWatcher(query);

            logWatcher.EventRecordWritten += new EventHandler<EventRecordWrittenEventArgs>(LogWatcher_EventRecordWritten);

            try
            {
                logWatcher.Enabled = true;
            }
            catch (EventLogException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }

        private static void LogWatcher_EventRecordWritten(object sender, EventRecordWrittenEventArgs e)
        {
            try
            {
                var time = e.EventRecord.TimeCreated;
                var id = e.EventRecord.Id;
                var logname = e.EventRecord.LogName;
                var level = e.EventRecord.Level;
                var providerName = e.EventRecord.ProviderName; // 来源字段 
                var task = e.EventRecord.TaskDisplayName == null ? "" : e.EventRecord.TaskDisplayName;
                var opCode = e.EventRecord.OpcodeDisplayName;
                var mname = e.EventRecord.MachineName;
                var xml = "";// e.EventRecord.ToXml();
                Console.WriteLine($@"【SourceMsg】:time:{time}, {id}, providerName:{providerName} logname:{logname}, level:{level}, task:{task}, OpcodeDisplayName:{opCode}, MachineName={mname},{xml}");
                // 检查事件来源是否为指定
                if (level == 2)
                {
                    if (CanSendNotification())
                    {
                        SendWeChatNotification(providerName, time.ToString());
                    }
                    else
                    {
                        Console.WriteLine("【TargetMsg】:超过每天和每分钟限额（防止微信封）.");
                    }
                    Console.WriteLine("");


                }
            }
            catch (Exception ee) {
                Console.WriteLine(ee.Message);
            }
            finally
            {

            }
            
        }
        private static bool CanSendNotification()
        {
            if (!File.Exists(LogFilePath))
            {
                File.WriteAllText(LogFilePath, JsonConvert.SerializeObject(new List<NotificationLog>()));
            }

            var logs = JsonConvert.DeserializeObject<List<NotificationLog>>(File.ReadAllText(LogFilePath));
            var now = DateTime.UtcNow;

            // 日志超过24小时移除
            logs.RemoveAll(log => now - log.Timestamp > TimeSpan.FromDays(1));

            // 检查每天是否超量
            if (logs.Count >= MaxNotificationsPerDay)
            {
                return false;
            }

            // 检查每分钟的量值
            if (logs.Any(log => now - log.Timestamp < NotificationInterval))
            {
                return false;
            }

            // 记录当前推送日志
            logs.Add(new NotificationLog { Timestamp = now });
            File.WriteAllText(LogFilePath, JsonConvert.SerializeObject(logs));

            return true;
        }
        private static async Task SendWeChatNotification(string operatype, string operacnt)
        {
            string accessToken = await GetAccessTokenAsync();
            Console.WriteLine("Access Token: " + accessToken);
            var url = $"https://api.weixin.qq.com/cgi-bin/message/template/send?access_token={accessToken}";

            var payload = new
            {
                touser = SendToOpenId,
                template_id = SendToTemplateId,
                data = new
                {
                    operatype = new { value = operatype, color = "#173177" },
                    operacnt = new { value = operacnt, color = "#173177" }
                }
            };

            var jsonPayload = JsonConvert.SerializeObject(payload);
            var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

            using (var client = new HttpClient())
            {
                var response = await client.PostAsync(url, content);
                var responseString = await response.Content.ReadAsStringAsync();
                Console.WriteLine("【Message】:WeChat Notification Response: " + responseString);
            }
        }
    }
}