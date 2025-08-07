using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Text;
using System.Xml.Linq;

namespace EZ_Get_ProjectProgress
{
    /*
     * 先GET EZTeamwork TMProjectService 取得所有專案
     * 再GET EZTeamwork TMTaskService 取得所有專案底下的所有任務、進度
     * 透過EZTeamwork專案的CRM交易編號，逐步更新CRM商機的專案進度
     * 首先，第一次PUT CRM交易的 EZTeamwork專案進度，將內容全部清除
     * 若清除成功，代表此專案存在
     * 第二次PUT 將EZTeamwork專案的任務、進度 更新。
     * 
     * 排程位置：192.168.100.212 的 C:\Scheduling\EZTeamwork\
     * 排程時間：每天00:00、12:00各執行一次
     */
    class Program
    {

        private static readonly HttpClientHandler handler = new HttpClientHandler
        {
            ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true // 忽略 SSL 驗證
        };
        public static readonly HttpClient Client = new HttpClient(handler);

        static async Task Main(string[] args)
        {
            List<ErrorLog> errorLogs = new List<ErrorLog>();

            try
            {
                DataTable dtNewProject = new DataTable();
                dtNewProject.Columns.Add("ProjID");
                dtNewProject.Columns.Add("ProjName");
                dtNewProject.Columns.Add("ProjObjKey");
                dtNewProject.Columns.Add("WBSCode");
                dtNewProject.Columns.Add("TaskName");
                dtNewProject.Columns.Add("Performance");
                dtNewProject.Columns.Add("CRMNo");

                string baseUrl = "https://ezteamwork.jmg.com.tw:18443/services/";
                //本機
                //string token = "6FD46F2011A65D81C7CA062D02DFF16467FB5C637A935FC8BCBE8B4A69FBA204930829F4929EDD5332A064FF8164468134101177FF44605F046999D9CDA458FDAE88EA99E93AB1DB1A250116DC5BB58A173D08C26BB7D7AA";
                //正式環境.212
                string token = "817E88150EA4CE8CB8ED924084676687BB3F1360572516FD8FC76DAB8A9BA2062582CFA368650EEFCEFBB8945C39180B4BBDC0B8EF62265FE3DE5B9C0233B7032BDA7D9F0CDF56C11EC72024BF60C972EADC8AE4F28E88CD";

                // Step 1: 查詢專案清單   (排除CRM交易編號是空的專案 AND convert(nvarchar(max),sda0) != '')
                string projectSoap = $@"<?xml version=""1.0"" encoding=""utf-8""?>
                                    <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
                                                   xmlns:sch=""http://soap.ws.tm.ctu.com"">
                                        <soap:Body>
                                            <sch:findProjectsTM75>
                                                <sch:securityToken>{token}</sch:securityToken>
                                                <sch:extraWhere>AND convert(nvarchar(max),sda0) != ''</sch:extraWhere>
                                            </sch:findProjectsTM75>
                                        </soap:Body>
                                    </soap:Envelope>";

                var projectRequest = new HttpRequestMessage(HttpMethod.Post, baseUrl + "TMProjectService");
                projectRequest.Headers.Add("SOAPAction", "findProjectsTM75");
                projectRequest.Content = new StringContent(projectSoap, Encoding.UTF8, "text/xml");

                var projectResponse = await Client.SendAsync(projectRequest);
                string projectXml = await projectResponse.Content.ReadAsStringAsync();

                //判斷有無資料
                if (projectXml.Contains("<ns:return xsi:nil=\"true\""))
                {
                    throw new Exception("EZ Teamwork查無資料，請確認securityToken");
                }

                // Step 2: 解析專案資料
                XDocument doc = XDocument.Parse(projectXml);
                var projects = doc.Descendants() // 取得所有後代元素
                                  .Where(e => e.Name.LocalName == "return") // 找到所有本地名稱為 "return" 的元素
                                  .Select(p => new // p 代表一個 "return" 元素
                                  {
                                      ProjID = p.Elements().FirstOrDefault(x => x.Name.LocalName == "projID")?.Value ?? "",
                                      ProjName = p.Elements().FirstOrDefault(x => x.Name.LocalName == "projName")?.Value ?? "",
                                      ProjObjKey = p.Elements().FirstOrDefault(x => x.Name.LocalName == "projObjKey")?.Value ?? "",
                                      CRMNo = p.Elements().FirstOrDefault(x => x.Name.LocalName == "sda0")?.Value ?? ""
                                  })
                                  .ToList();

                // Step 3: 對每個專案查詢任務清單
                foreach (var project in projects)
                {
                    if (string.IsNullOrWhiteSpace(project.ProjObjKey))
                    {
                        continue;
                    }

                    //排除ProjID是空的專案
                    if (string.IsNullOrWhiteSpace(project.ProjID))
                    {
                        continue;
                    }

                    string taskSoap = $@"<?xml version=""1.0"" encoding=""utf-8""?>
                                    <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
                                                   xmlns:sch=""http://soap.ws.tm.ctu.com"">
                                        <soap:Body>
                                            <sch:findActiveTasksByProjObjKey>
                                                <sch:projObjKey>{project.ProjObjKey}</sch:projObjKey>
                                                <sch:securityToken>{token}</sch:securityToken>
                                            </sch:findActiveTasksByProjObjKey>
                                        </soap:Body>
                                    </soap:Envelope>";

                    var taskRequest = new HttpRequestMessage(HttpMethod.Post, baseUrl + "TMTaskService");
                    taskRequest.Headers.Add("SOAPAction", "findActiveTasksByProjObjKey");
                    taskRequest.Content = new StringContent(taskSoap, Encoding.UTF8, "text/xml");

                    var taskResponse = await Client.SendAsync(taskRequest);
                    string taskXml = await taskResponse.Content.ReadAsStringAsync();

                    // 解析任務 XML
                    XDocument taskDoc = XDocument.Parse(taskXml);
                    var tasks = taskDoc.Descendants()
                                       .Where(e => e.Name.LocalName == "return")
                                       .Select(t => new
                                       {
                                           WBSCode = t.Elements().FirstOrDefault(x => x.Name.LocalName == "WBSCode")?.Value ?? "N/A",
                                           TaskName = t.Elements().FirstOrDefault(x => x.Name.LocalName == "taskName")?.Value
                                                                                                                      ?.Replace("\r\n", "")  // 移除 Windows 換行
                                                                                                                      ?.Replace("\n", "")    // 移除 Unix/Linux/macOS 換行
                                                                                                                      ?.Replace("\r", "")    // 移除舊版 Mac 換行
                                                                                                                      ?? "N/A",
                                           Performance = t.Elements().FirstOrDefault(x => x.Name.LocalName == "performancePercentage")?.Value ?? "N/A"
                                       })
                                       .Where(t => t.WBSCode != null && !t.WBSCode.Contains(".")) // 只保留沒有點的（即最上層）
                                       .ToList();

                    foreach (var task in tasks)
                    {
                        dtNewProject.Rows.Add(project.ProjID, project.ProjName, project.ProjObjKey, task.WBSCode, task.TaskName, task.Performance, project.CRMNo);
                    }
                }

                //取得CRM token
                string accessToken = "";
                if (dtNewProject.Rows.Count > 0)
                {
                    var url = "https://accounts.zoho.com/oauth/v2/token";
                    var content = new FormUrlEncodedContent(new[]
                    {
                        //測試
                        //new KeyValuePair<string, string>("refresh_token", "1000.836491c055ddd474a16adca4e11a9c53.41cfc8883a5c955835b27f910b9f6f2c"),
                        //new KeyValuePair<string, string>("client_id", "1000.CDKBG6X1YZ328YUZJGYR2C881GNJFH"),
                        //new KeyValuePair<string, string>("client_secret", "f249e3a69f6eb110638568c57448bf79ba3e1cf5ab"),
                        //正式
                        new KeyValuePair<string, string>("refresh_token", "1000.4a8cd29eea1dea8be07d7180e0e7b319.8c508dced8639df35ef3ee7fb47d2170"),
                        new KeyValuePair<string, string>("client_id", "1000.AFH6759VOUNZI90JZ72HCD70BFHVEC"),
                        new KeyValuePair<string, string>("client_secret", "60125a66cec3e8cc36560cdcff6082caea02c9d941"),
                        new KeyValuePair<string, string>("grant_type", "refresh_token")
                    });

                    HttpResponseMessage response = await Client.PostAsync(url, content);
                    string tokenContent = await response.Content.ReadAsStringAsync();
                    if (response.IsSuccessStatusCode && !tokenContent.Contains("error"))
                    {
                        var json = JObject.Parse(tokenContent);
                        accessToken = json["access_token"]?.ToString();
                    }
                    else
                    {
                        throw new Exception($"Failed to get access token: {tokenContent}");
                    }
                }
                else
                {
                    Environment.Exit(0);
                }

                //Distinct ProjID
                DataTable distinctProjectID = dtNewProject.DefaultView.ToTable(true, "ProjID", "CRMNo");
                foreach (DataRow drProjectID in distinctProjectID.Rows)
                {
                    string selProj = drProjectID["ProjID"]?.ToString() ?? "";
                    string selCRMNo = drProjectID["CRMNo"]?.ToString() ?? "";

                    //PUT 先清空原本的專案狀態
                    var deleteBody = new
                    {
                        data = new[]
                        {
                            new
                            {
                                EXT_CRM_No = selCRMNo,
                                EZTeamwork = new object[] { } // 空陣列
                            }
                        }
                    };

                    var deleteJson = JsonConvert.SerializeObject(deleteBody);
                    var deleteRequest = new HttpRequestMessage(HttpMethod.Put, "https://www.zohoapis.com/crm/v7/Deals");
                    deleteRequest.Headers.Authorization = new AuthenticationHeaderValue("Zoho-oauthtoken", accessToken);
                    deleteRequest.Headers.Add("X-EXTERNAL", "Deals.EXT_CRM_No");
                    deleteRequest.Content = new StringContent(deleteJson, Encoding.UTF8, "application/json");

                    var deleteResponse = await Client.SendAsync(deleteRequest);
                    string deleteContent = await deleteResponse.Content.ReadAsStringAsync();
                    if (!deleteResponse.IsSuccessStatusCode)
                    {
                        if (!deleteContent.Contains("the external id given seems to be invalid"))
                        {
                            errorLogs.Add(new ErrorLog
                            {
                                Timestamp = DateTime.Now,
                                ProjectID = selProj,
                                Action = "清除 EZ Teamwork 專案狀態",
                                Message = deleteContent
                            });
                        }
                        continue;
                    }

                    //PUT CRM交易的專案狀態 (透過交易ID)
                    List<Dictionary<string, string>> ezTeamworkList = new List<Dictionary<string, string>>();
                    DataRow[] drProj = dtNewProject.Select($"ProjID = '{selProj}'");
                    foreach (DataRow dr in drProj)
                    {
                        ezTeamworkList.Add(new Dictionary<string, string>
                        {
                            { "Stage", dr["WBSCode"].ToString() },
                            { "Task", dr["TaskName"].ToString() },
                            { "Status", double.Parse(dr["Performance"].ToString()).ToString("0.#") + "%" }
                        });
                    }
                    // 組整個 JSON 資料結構
                    var updateBody = new
                    {
                        data = new[]
                        {
                            new
                            {
                                EXT_CRM_No = selCRMNo,
                                SAP_No = selProj,   //修改交易的SAP專案編號
                                EZTeamwork = ezTeamworkList
                            }
                        }
                    };
                    var updateJson = JsonConvert.SerializeObject(updateBody);
                    var updateRequest = new HttpRequestMessage(HttpMethod.Put, "https://www.zohoapis.com/crm/v7/Deals");
                    updateRequest.Headers.Authorization = new AuthenticationHeaderValue("Zoho-oauthtoken", accessToken);
                    updateRequest.Headers.Add("X-EXTERNAL", "Deals.EXT_CRM_No");
                    updateRequest.Content = new StringContent(updateJson, Encoding.UTF8, "application/json");

                    var updateResponse = await Client.SendAsync(updateRequest);
                    string updateContent = await updateResponse.Content.ReadAsStringAsync();
                    if (!updateResponse.IsSuccessStatusCode)
                    {
                        errorLogs.Add(new ErrorLog
                        {
                            Timestamp = DateTime.Now,
                            ProjectID = selProj,
                            Action = "修改 EZ Teamwork 專案狀態",
                            Message = updateContent
                        });
                    }

                }
            }
            catch (Exception ex)
            {
                errorLogs.Add(new ErrorLog
                {
                    Timestamp = DateTime.Now,
                    ProjectID = "無",
                    Action = "例外錯誤",
                    Message = ex.Message
                });
            }
            finally
            {
                if (errorLogs.Count > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("錯誤日誌如下：\n");
                    foreach (var log in errorLogs)
                    {
                        sb.AppendLine($"{log.Timestamp:yyyy-MM-dd HH:mm:ss} | {log.ProjectID} | {log.Action} | {log.Message}");
                    }
                    //寄信
                    ErrSendEmail("【排程】EZ Teamwork專案狀態更新異常", sb.ToString());
                }
                Environment.Exit(0);
            }
        }

        //異常發信
        public static void ErrSendEmail(string subject, string body)
        {
            // 設定 SMTP 伺服器的資訊
            SmtpClient smtpClient = new SmtpClient("smtp.office365.com") // 使用你 SMTP 伺服器的地址
            {
                Port = 587, // SMTP 郵件伺服器的端口，通常是 587 或 465
                Credentials = new NetworkCredential("jiinming@mail.jmg.com.tw", "Ji2249mI0277"), // 用來登入的帳戶和密碼
                EnableSsl = true // 启用 SSL 加密
            };
            // 創建郵件訊息
            MailMessage mailMessage = new MailMessage
            {
                From = new MailAddress("jiinming@mail.jmg.com.tw"), // 發件人地址
                Subject = subject, // 郵件標題
                Body = body, // 郵件內容
                IsBodyHtml = false // 設置郵件內容是否為 HTML 格式
            };
            // 添加收件人
            mailMessage.To.Add("J1@jmg.com.tw"); // 分割並加入每個收件人
            // 發送郵件
            smtpClient.Send(mailMessage);
        }
    }

    public class ErrorLog
    {
        public DateTime Timestamp { get; set; }
        public string ProjectID { get; set; }
        public string Action { get; set; }
        public string Message { get; set; }
    }
}
