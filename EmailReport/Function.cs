using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Threading.Tasks;
using Amazon;
using Amazon.Lambda.Core;
using Amazon.S3;
using Amazon.S3.Model;
using Amazon.SimpleEmail;
using Amazon.SimpleEmail.Model;
using MimeKit;
using OfficeOpenXml;
using OfficeOpenXml.Style;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace EmailReport
{
    public class FunctionInput
    {
        public string inputYear { get; set; }
        public string inputRunDate { get; set; }
    }

    public class Function
    {
        static int year = 0;
        static DateTime dataDate = DateTime.MinValue;
        static RegionEndpoint region = RegionEndpoint.APSoutheast1;
        public async Task<string> FunctionHandler(FunctionInput input, ILambdaContext context)
        {
            if (string.IsNullOrEmpty(input.inputYear))
            {
                year = DateTime.Now.Year;
            }
            else
            {
                year = int.Parse(input.inputYear);
            }

            if (!string.IsNullOrEmpty(input.inputRunDate))
            {
                dataDate = DateTime.ParseExact(input.inputRunDate, "yyyy-MM-dd", null);
            }

            var dataSummary = LoadSummary();
            var send = StartSendEmail();

            Task.WaitAll(dataSummary, send);

            var date = dataSummary.Result.Keys.Max(m => m);
            var dateThai = $"{date.ToString("dd/MM/")}{date.Year + 543}";
            var dateFile = date.ToString("yyyyMMdd");


            var sendTime = $"{send.Result.ToString("dd/MM/")}{send.Result.Year + 543} " + send.Result.ToString("HH:mm:ss");


            if (date != DateTime.MinValue)
            {
                var dataBounce = LoadBounce(dataDate);
                var branch = LoadBranch();
                var streamSummary = ReportSummary(dataSummary.Result, dateFile);

                var streamBounce = ReportBounce(await dataBounce, branch, dateFile);
                
                await Task.WhenAll(streamBounce, streamSummary);

                using var messageStream = new MemoryStream();
                var message = new MimeMessage();

                message.Subject = $"รายงานผลการส่ง Email งาน Daily - Equity รอบงาน {dateThai}";

                message.From.Add(InternetAddress.Parse(Utility.Env("EMAIL_FROM")));
                message.To.Add(InternetAddress.Parse("s92347@scb.co.th"));
                var cc = Utility.Env("EMAIL_CC");
                if (!string.IsNullOrEmpty(cc))
                {
                    foreach (var address in cc.Split(','))
                    {
                        message.Cc.Add(InternetAddress.Parse(address));
                    }

                }

                var body = new BodyBuilder();
                var sumarry = dataSummary.Result[date.Date];
                var unknow = dataBounce.Result.Where(w=>w.Remark == "CO : No Response").Count();
                body.HtmlBody = Html(dateThai, sendTime, sumarry.Success, sumarry.Fail - unknow, unknow);
                
                body.Attachments.Add($"report_summary_{dateFile}.xlsx", streamSummary.Result);
                body.Attachments.Add($"report_fail_{dateFile}.xlsx", streamBounce.Result);
   
                message.Body = body.ToMessageBody();
                message.WriteTo(messageStream);

                var email = new SendRawEmailRequest()
                {
                    RawMessage = new RawMessage() { Data = messageStream }
                };

                var mailClient = new AmazonSimpleEmailServiceClient(region);
                var result = await mailClient.SendRawEmailAsync(email);

                return "success";
            }
            else
            {
                return "no data";
            }

            
        }

        public static async Task<MemoryStream> ReportBounce(List<Bounce> data, List<Branch> branch, string dateText)
        {
            var sep = Path.DirectorySeparatorChar;
            var excel = new FileInfo($"{Environment.CurrentDirectory}{sep}bounce.xlsx");

            var stream = new MemoryStream();
            using (var excelPackage = new ExcelPackage(excel))
            {
                var sheet = 0;
                var cells = excelPackage.Workbook.Worksheets[sheet].Cells;
                int i = 1;

                foreach (var item in data)
                {
                    i++;
                    var bname = branch.Where(w => w.Team.Contains(item.Team)).FirstOrDefault();
                    cells["A" + i].Value = item.Date.ToString("dd/MM/yyyy");
                    cells["B" + i].Value = item.Account;
                    cells["C" + i].Value = item.Team != "NA" ? item.Team : "";
                    cells["D" + i].Value = bname != null ? bname.Name : "";
                    cells["E" + i].Value = item.Name;
                    cells["F" + i].Value = item.Email;
                    cells["G" + i].Value = item.Remark;
                    cells["H" + i].Value = item.DocType;
                    cells["I" + i].Value = item.AoId;
                    cells["J" + i].Value = item.AoName;

                }

                if (i != 1)
                {
                    var range = $"A2:J{i}";
                    cells[range].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    cells[range].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    cells[range].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    cells[range].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                }

                sheet++;


                foreach (var b in branch)
                {
                    i = 1;
                    cells = excelPackage.Workbook.Worksheets[sheet].Cells;

                    foreach (var item in data.Where(w => b.Team.Contains(w.Team)))
                    {
                        i++;
                        cells["A" + i].Value = item.Date.ToString("dd/MM/yyyy"); ;
                        cells["B" + i].Value = item.Account;
                        cells["C" + i].Value = item.Name;
                        cells["D" + i].Value = item.Email;
                        cells["E" + i].Value = item.Remark;
                        cells["F" + i].Value = item.DocType;
                        cells["G" + i].Value = item.AoId;
                        cells["H" + i].Value = item.AoName;

                    }
                    if (i != 1)
                    {
                        var range = $"A2:H{i}";
                        cells[range].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        cells[range].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        cells[range].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        cells[range].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    }


                    sheet++;
                }

                excelPackage.SaveAs(stream);

                await WriteToS3($"report_fail_{dateText}.xlsx", stream);
                stream.Seek(0, SeekOrigin.Begin);

                return stream;
            }
        }

        public static async Task<MemoryStream> ReportSummary(Dictionary<DateTime, Summary> data, string dateText)
        {

            var sep = Path.DirectorySeparatorChar;
            var excel = new FileInfo($"{Environment.CurrentDirectory}{sep}summary.xlsx");

            var stream = new MemoryStream();
            using (var excelPackage = new ExcelPackage(excel))
            {
                var colorSuccess = System.Drawing.Color.FromArgb(204, 153, 255);
                var colorFail = System.Drawing.Color.FromArgb(244, 176, 132);

                int i = 5;
                var cells = excelPackage.Workbook.Worksheets[0].Cells;
                DateTime start = new DateTime(year, 1, 1);
                DateTime end = new DateTime(year + 1, 1, 1);
                var month = 1;
                do
                {
                    if (month != start.Month)
                    {
                        month++;

                        if (month == 3 && year % 4 != 0)
                        {
                            cells["A67"].Value = "";
                            i += 4;
                        }
                        else
                        {
                            i += 3;
                        }
                    }
                    var fillColor = false;

                    cells["A" + i].Value = start.ToString("dd/MM/yyyy");
                    if (data.ContainsKey(start))
                    {

                        if (data[start].DocType == "SDC")
                        {
                            fillColor = true;

                            cells["E" + i].Style.Fill.BackgroundColor.SetColor(colorSuccess);
                            cells["E" + i].Style.Fill.BackgroundColor.Tint = 0;

                            cells["F" + i].Value = data[start].Success;
                            cells["F" + i].Style.Fill.BackgroundColor.SetColor(colorSuccess);
                            cells["F" + i].Style.Fill.BackgroundColor.Tint = 0;

                            cells["G" + i].Value = data[start].Fail;
                            cells["G" + i].Style.Fill.BackgroundColor.Tint = 0;
                            cells["G" + i].Style.Fill.BackgroundColor.SetColor(colorFail);
                        }
                    }

                    if (fillColor)
                    {

                        cells["B" + i].Style.Fill.BackgroundColor.SetColor(colorSuccess);
                        cells["B" + i].Style.Fill.BackgroundColor.Tint = 0;
                        cells["C" + i].Style.Fill.BackgroundColor.SetColor(colorSuccess);
                        cells["C" + i].Style.Fill.BackgroundColor.Tint = 0;
                        cells["D" + i].Style.Fill.BackgroundColor.SetColor(colorFail);
                        cells["D" + i].Style.Fill.BackgroundColor.Tint = 0;
                    }

                    i++;
                    start = start.AddDays(1);
                }
                while (start < end);

                excelPackage.SaveAs(stream);
                await WriteToS3($"report_summary_{dateText}.xlsx", stream);
                stream.Seek(0, SeekOrigin.Begin);

                return stream;
            }
        }

        public static async Task<List<Bounce>> LoadBounce(DateTime dataDate)
        {
            var account = "";
            var doc = "";
            var step = "";
            var status = "";
            var noteRemark = "";
            try
            {

                var result = new List<Bounce>();
                var sql = $@"WITH data 
                                             AS (SELECT account,
                                                        document_type,
                                                        step_name,
                                                        status,
                                                        message_payload,
                                                        remark,
                                                        data_date,
                                                        Row_number() over ( 
                                                            PARTITION BY account 
                                                            ORDER BY id DESC) AS num 
                                                 FROM   public.job_step_execution
                                                 WHERE (step_name = 'SENDING_EMAIL' OR step_name = 'SENT_EMAIL' OR step_name = 'STATUS_EMAIL') 
                                                 AND document_type = 'SDC' AND product = 'EQUITY' AND data_date = '{dataDate}')
                                        SELECT * 
                                        FROM   data 
                                        WHERE  num = 1
                                        ORDER BY document_type";

                using var conn = await Utility.CreateConnection();

                var failList = new Dictionary<string, string>();
                var accounts = new System.Text.StringBuilder();
                using (var reader = await conn.ExecuteReaderAsync(sql))
                {
                    while (reader.Read())
                    {

                        step = reader.GetString(2);
                        status = reader.GetString(3);
                        noteRemark = reader.GetString(5);

                        var check = (step == "SENT_EMAIL" && status == "SUCCESS") ? true : false;

                        if (!check)
                        {
                            account = reader.GetString(0);
                            doc = reader.GetString(1);
                            if (
                                step == "SENDING_EMAIL" ||
                                step == "STATUS_EMAIL" ||
                                (step == "SENT_EMAIL" && status != "SUCCESS" ))
                            {
                                var root = JsonDocument.Parse(reader.GetString(4)).RootElement;

                                if (step == "SENT_EMAIL" && status != "SUCCESS") {
                                    var sqlQuery =  $@"SELECT * FROM public.job_step_execution WHERE account = '{account}' AND step_name = 'SENDING_EMAIL' LIMIT 1";
                                    var sqlResult = await conn.ExecuteScalarAsync(sqlQuery);
                                    var edok = sqlResult.ToString();                
                                    Console.WriteLine(edok);
                                }

                                var documentMeta = root.GetProperty("documentMeta");
                                var deliveryInfo = documentMeta.GetProperty("deliveryInfo");
                                var marketingInfo = documentMeta.GetProperty("marketingInfo");
                                var bounce = new Bounce();
                                bounce.Account = deliveryInfo.GetProperty("accountNumberFormat").GetString();
                                bounce.Name = deliveryInfo.GetProperty("customerFullName").GetString();
                                bounce.Email = deliveryInfo.GetProperty("customerEmail").GetString();
                                bounce.AoName = marketingInfo.GetProperty("marketingName").GetString();
                                bounce.AoId = marketingInfo.GetProperty("marketingId").GetString();
                                bounce.Team = marketingInfo.GetProperty("teamId").GetString() ?? "NA";
                                bounce.Remark = status == "CO" ? "CO : No Response" : status;

                                if (status == "CO_SUCCESS") {
                                    bounce.Remark = "CO : Response Success " + noteRemark;
                                } else if (status == "CO_BOUNCE") {
                                    bounce.Remark = "CO : Response FAILED " + noteRemark;
                                }

                                bounce.DocType = doc;
                                bounce.Date = reader.GetDateTime(6);

                                result.Add(bounce);

                            }
                            else
                            {
                                failList.Add(account + doc, reader.GetString(5));
                                accounts.Append($",'{account}'");
                            }

                        }
                    }
                }

                if (accounts.Length != 0)
                {
                    sql = $@"SELECT 
                                      data_date,
                                      document_type,
                                      account,
                                      message_payload#>>'{{documentMeta,deliveryInfo,accountNumberFormat}}' as account,
                                      message_payload#>>'{{documentMeta,deliveryInfo,customerFullName}}' as name,
                                      message_payload#>>'{{documentMeta,deliveryInfo,customerEmail}}' as email,
                                      message_payload#>>'{{documentMeta,marketingInfo,marketingName}}' as mkt,
                                      message_payload#>>'{{documentMeta,marketingInfo,marketingId}}' as mkt_id,
                                      message_payload#>>'{{documentMeta,marketingInfo,teamId}}' as team_id
                                    FROM  public.job_step_execution 
                                    WHERE step_name = 'SENDING_EMAIL' AND document_type = 'SDC' AND product = 'EQUITY' AND account IN ({accounts.Remove(0, 1)})";

                    using (var reader = await conn.ExecuteReaderAsync(sql))
                    {
                        while (reader.Read())
                        {
                            account = reader.GetString(2);
                            doc = reader.GetString(1);
                            if (failList.TryGetValue(account + doc, out var fail))
                            {

                                result.Add(new Bounce
                                {
                                    Date = reader.GetDateTime(0),
                                    DocType = doc,
                                    Remark = fail,
                                    Account = reader.GetString(3),
                                    Name = reader.GetString(4),
                                    Email = reader.GetString(5),
                                    AoName = !reader.IsDBNull(6) ? reader.GetString(6) : "",
                                    AoId = !reader.IsDBNull(7) ? reader.GetString(7) : "",
                                    Team = !reader.IsDBNull(8) ? reader.GetString(8).Trim() : "NA",
                                });
                            }

                        }
                    }
                }

                return result;
            }
            catch
            {
                Console.WriteLine(account);
                Console.WriteLine(step);
                Console.WriteLine(status);
                throw;
            }
            
        }

        public static async Task<Dictionary<DateTime, Summary>> LoadSummary()
        {
            var start = new DateTime(year, 1, 1).ToString("MM-dd-yyyy");
            var end = new DateTime(year, 12, 31).ToString("MM-dd-yyyy");
            var result = new Dictionary<DateTime, Summary>();
            var sql = $@"SELECT 
                                      data_date,
                                      document_type,
                                      success_num,
                                      fail_num
                                    FROM 
                                      public.email_status_summary
                                    WHERE
                                      document_type = 'SDC' ";

            using (var conn = await Utility.CreateConnection())
            using (var reader = await conn.ExecuteReaderAsync(sql))
            {
                while (reader.Read())
                {
                    var date = reader.GetDateTime(0);
                    var doc = reader.GetString(1);
                    var success = reader.GetInt32(2);
                    var fail = reader.GetInt32(3);
                    result[date] = new Summary { DocType = doc, Success = success, Fail = fail };
                }
            }
            return result;
        }

        private static List<Branch> LoadBranch()
        {
            var sep = Path.DirectorySeparatorChar;
            var excel = new FileInfo($"{Environment.CurrentDirectory}{sep}branch.xlsx");
            var result = new List<Branch>();
            using (var excelPackage = new ExcelPackage(excel))
            {
                var cells = excelPackage.Workbook.Worksheets[0].Cells;
                for (int i = 3; ; i++)
                {
                    if (cells["A" + i].Value != null && !string.IsNullOrEmpty(cells["A" + i].Text))
                    {
                        var b = result.FirstOrDefault(a => a.Name == cells["C" + i].Text.Trim());
   
                        if (b != null)
                        {
                            b.Team += "," + cells["A" + i].Text.Trim();
                        }
                        else
                        {
                            result.Add( new Branch { Name = cells["C" + i].Text.Trim(), Team = cells["A" + i].Text.Trim() });
                        }

                    }
                    else
                    {
                        break;
                    }
                }
            }

            return result;
        }

        private static async Task<DateTime> StartSendEmail()
        {
            var sql = $"SELECT create_datetime FROM public.job_step_execution WHERE step_name='SENDING_EMAIL' AND document_type = 'SDC' AND product = 'EQUITY' ORDER BY id LIMIT 1";
            using (var conn = await Utility.CreateConnection())
            {
                var result = await conn.ExecuteScalarAsync(sql);
                return result != null ? (DateTime)result : DateTime.MinValue;
            }
        }

        public static Task WriteToS3(string file, MemoryStream data)
        {
            var client = new AmazonS3Client(region);
            var request = new PutObjectRequest();
            request.BucketName = Utility.Env("BUCKET");
            request.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            request.Key = $"{Utility.Env("PATH")}{file}";

            var clone = new MemoryStream();

            data.WriteTo(clone);
            clone.Seek(0, SeekOrigin.Begin);

            request.InputStream = clone;

            return client.PutObjectAsync(request);
        }

        public async Task<bool> ExitInS3(string file)
        {
            try
            {
                var client = new AmazonS3Client(region);
                var request = new GetObjectMetadataRequest();
                request.BucketName = Utility.Env("BUCKET");
                request.Key = $"{Utility.Env("PATH")}{file}";
                Console.WriteLine(request.Key);
                var result = await client.GetObjectMetadataAsync(request);
                return true;
            }
            catch
            {
                return false;
            }

        }



        private static string Html(string date, string sendTime, int sucess, int fail, int unknow)
        {
            var template = $@"
<!DOCTYPE html><html><head>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>
<style>

</style>
</head>
<body>

<table style='border-collapse: collapse;border-spacing: 0;width: 100%;'>
<tr>
	<td style='font-family: Tahoma,Arial, sans-serif;font-size: 14px;font-weight: bold;padding-top:20px; padding-bottom:20px;'>
	งาน Daily - Equity (รอบงาน {date}) 
	</td>
</tr>
</table>
<br>


<div style='overflow-x:auto;'>
  <table style='border-collapse: collapse;border-spacing: 0;width: 100%;border: 1px solid #ddd;'>
    <tr>
      <th bgcolor='#f2f2f2' style='text-align: left;padding: 8px;border: 1px solid #ddd;font-family: Tahoma,Arial, sans-serif;font-size: 12px;'>วันที่ส่ง</th>
      <th bgcolor='#f2f2f2' style='text-align: left;padding: 8px;border: 1px solid #ddd;font-family: Tahoma,Arial, sans-serif;font-size: 12px;'>ยอดทั้งหมด (ราย)</th>
      <th bgcolor='#f2f2f2' style='text-align: left;padding: 8px;border: 1px solid #ddd;font-family: Tahoma,Arial, sans-serif;font-size: 12px;'>ส่งสำเร็จ (ราย)</th>
      <th bgcolor='#f2f2f2' style='text-align: left;padding: 8px;border: 1px solid #ddd;font-family: Tahoma,Arial, sans-serif;font-size: 12px;'>ส่งไม่สำเร็จ (ราย) </th>
      <th bgcolor='#f2f2f2' style='text-align: left;padding: 8px;border: 1px solid #ddd;font-family: Tahoma,Arial, sans-serif;font-size: 12px;'>ไม่ทราบผล (ราย) </th>
      <th bgcolor='#f2f2f2' style='text-align: left;padding: 8px;border: 1px solid #ddd;font-family: Tahoma,Arial, sans-serif;font-size: 12px;'>งาน </th>
    </tr>
    <tr>
                      <td style='text-align: left;padding: 8px;border: 1px solid #ddd;font-family: Tahoma,Arial, sans-serif;font-size: 12px;'>{sendTime}</td>
                      <td style='text-align: left;padding: 8px;border: 1px solid #ddd;font-family: Tahoma,Arial, sans-serif;font-size: 12px;'>{(sucess + fail + unknow).ToString("#,0")}</td>
                      <td style='text-align: left;padding: 8px;border: 1px solid #ddd;font-family: Tahoma,Arial, sans-serif;font-size: 12px;'>{sucess.ToString("#,0")}</td>
                      <td style='text-align: left;padding: 8px;border: 1px solid #ddd;font-family: Tahoma,Arial, sans-serif;font-size: 12px;'>{fail.ToString("#,0")}</td>
                      <td style='text-align: left;padding: 8px;border: 1px solid #ddd;font-family: Tahoma,Arial, sans-serif;font-size: 12px;'>{unknow.ToString("#,0")}</td>
                      <td style='text-align: left;padding: 8px;border: 1px solid #ddd;font-family: Tahoma,Arial, sans-serif;font-size: 12px;'>SDC</td>
                    </tr>
  </table>
</div>

</body>
</html>

";
            return template;
        }

    }

    public struct Summary
    {
        public string DocType { get; set; }
        public int Success { get; set; }
        public int Fail { get; set; }
    }

    public struct Bounce
    {
        public DateTime Date { get; set; }
        public string DocType { get; set; }
        public string Remark { get; set; }
        public string Account { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string AoName { get; set; }
        public string AoId { get; set; }
        public string Team { get; set; }
    }

    public class Branch
    {
        public string Team { get; set; }
        public string Name { get; set; }
    }
}
