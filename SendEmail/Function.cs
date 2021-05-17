using Amazon;
using Amazon.Lambda.Core;
using Amazon.Lambda.SQSEvents;
using Amazon.S3;
using Amazon.S3.Model;
using Amazon.SecretsManager;
using Amazon.SimpleEmail;
using Amazon.SimpleEmail.Model;
using MimeKit;
using System;
using System.Data.Common;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Threading.Tasks;


// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace SendEmail
{
    public class Function
    {

        static RegionEndpoint region = RegionEndpoint.APSoutheast1;
        static JsonSerializerOptions serializerOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        };

        static MailboxAddress from = new MailboxAddress(Environment.GetEnvironmentVariable("EMAIL_NAME"), Environment.GetEnvironmentVariable("EMAIL_FROM"));
        static string cc = Environment.GetEnvironmentVariable("EMAIL_CC");
        public async Task FunctionHandler(SQSEvent sqsEvent, ILambdaContext context)
        {
            foreach (var item in sqsEvent.Records)
            {
                var dup = await CheckDuplicate(item.MessageId);
                if(!dup)
                {
                    await Process(item);
                }

            }
        }

        public async Task Process(SQSEvent.SQSMessage sqs)
        {
            SendMail data = null;
            var date = DateTime.MinValue;
            try
            {
                data = JsonSerializer.Deserialize<SendMail>(sqs.Body, serializerOptions);
                date = DateTime.ParseExact(data.DocumentMeta.DataDate, "yyyy-MM-dd", null);
                using (var mailClient = new AmazonSimpleEmailServiceClient(region))
                using (var messageStream = new MemoryStream())
                {

                    var message = BuildMessage(data.DocumentMeta, date);
                    
                    var body = await BuildBody(data);
                    
                    message.Body = body.ToMessageBody();
                    message.WriteTo(messageStream);

                    var email = new SendRawEmailRequest()
                    {
                        RawMessage = new RawMessage() { Data = messageStream }
                    };

                    var result = await mailClient.SendRawEmailAsync(email);
 
                    await LogToDB(sqs.Body, data, date, sqs.MessageId, null, result.MessageId);

                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                await LogToDB(sqs.Body, data, date, sqs.MessageId, ex.Message, null);
            }

        }

        private static async Task<bool> CheckDuplicate(string msgId)
        {
            var sql = $@"SELECT 1 FROM public.job_step_execution WHERE message_id = '{msgId}' AND status = 'SUCCESS' LIMIT 1";
            using (var conn = await Utility.CreateConnection())
            {
                var result = await conn.ExecuteScalarAsync(sql);
                return result != null;
            }
        }

        private static MimeMessage BuildMessage(DocumentMeta data, DateTime date)
        {
            var message = new MimeMessage();
            var dateText = date.ToString("dd/MM/yyyy");

            if (data.DocumentType == "SDC") {
                message.Subject = $"Daily Confirmation trade date {dateText} and or Deposit/Receive Slip and or Cash Withdrawal Slip of {data.DeliveryInfo.AccountNumberFormat} {data.DeliveryInfo.CustomerFullName}";
            } else if (data.DocumentType == "FX_PAYSLIP") {
                message.Subject = $"SCBS: Exchange Slip as of {dateText}";
            }


            message.Headers.Add("DataDate", dateText);
            message.Headers.Add("Product", data.Product);
            message.Headers.Add("DocumentType", data.DocumentType);
            message.Headers.Add("JobExecutionId", $"{data.JobExecutionId}");
            message.Headers.Add("AccountNumber", $"{data.AccountNumber}");

            if (data.ExecutionCutoffTime != null) {
                message.Headers.Add("ExecutionCutoffTime", data.ExecutionCutoffTime);
            }

            message.From.Add(from);
            message.To.Add(InternetAddress.Parse(data.DeliveryInfo.CustomerEmail));
            if (!string.IsNullOrEmpty(cc))
            {
                message.Cc.Add(InternetAddress.Parse(cc));
            }

            return message;
        }

        private static async Task<BodyBuilder> BuildBody(SendMail data)
        {
            var body = new BodyBuilder();
            if(data.DocumentMeta.DocumentType == "SDC")
            {
                body.HtmlBody = SDC(data.DocumentMeta);
            } else if (data.DocumentMeta.DocumentType == "FX_PAYSLIP") {
                body.HtmlBody = FXPAYSLIP(data.DocumentMeta);
            }

            var s3Client = new AmazonS3Client(region);

            var request = new GetObjectRequest();

            foreach (var item in data.Documents)
            {
                request.BucketName = item.BucketName;
                request.Key = item.Path;

                using (var response = await s3Client.GetObjectAsync(request))
                using (var strean = response.ResponseStream)
                {
                    var index = request.Key.LastIndexOf('/') + 1;
                    var fileName = request.Key.Substring(index, request.Key.Length - index);
                    body.Attachments.Add(fileName, strean);
                }
            }

            return body;
        }

        private static string FXPAYSLIP(DocumentMeta data) {
            var template = $@"<html>

<head> </head>

<body>
    <div dir='auto'>
        <blockquote type='cite'>
            <div dir='ltr'>
                <meta content='text/html; charset=utf-8'>
                <table border='0' cellpadding='0' cellspacing='0'>
                    <tbody>
                        <tr>
                            <td>
                                <div><span style='font-size:14.0pt'><span
                                            style='font-family:cordia new,sans-serif'>เรียน&nbsp;{data.DeliveryInfo.CustomerFullName}</span></span>
                                </div>
                                <div><span style='font-size:14.0pt;font-family:cordia new,sans-serif'>
                                        บริษัทหลักทรัพย์ ไทยพาณิชย์ จำกัด (“บริษัทฯ”)
                                        ขอนำส่งใบยืนยันการซื้อขาย/รายงานแสดงทรัพย์สิน/ใบกำกับภาษี และ/หรือ ใบรับฝาก
                                        และ/หรือ ใบส่งมอบเงิน และ/หรือ ใบแจ้งการแลกเปลี่ยนเงินตรา ในรูปแบบ PDF file
                                        ตามแนบ ท่านสามารถเปิดดูไฟล์ข้อมูลได้ด้วยรหัสผ่านของท่าน
                                        โดยรหัสผ่านของท่านคือ เลขประจำตัวประชาชน 6 หลักสุดท้าย (เช่น
                                        เลขประจำตัวประชาชน 1-1115-00240-43-6 รหัสผ่านคือ 240436) <br />
                                        เอกสารที่แนบมานี้จะอยู่ในรูปแบบ PDF ซึ่งเปิดดูได้ด้วยโปรแกรม Adobe Reader
                                        เวอร์ชั่น 6.0 ขึ้นไป หากท่านไม่สามารถเปิดเอกสารดังกล่าวได้
                                        สามารถดาวน์โหลดโปรแกรมได้ฟรีที่ http://www.adobe.com/products/acrobat/
                                        <br /> อีเมลฉบับนี้เป็นการแจ้งข้อมูลจากระบบโดยอัตโนมัติ กรุณาอย่าตอบกลับ
                                        หากท่านมีข้อสงสัยหรือต้องการสอบถามรายละเอียดเพิ่มเติม กรุณาติดต่อ
                                        ที่ฝ่ายลูกค้าสัมพันธ์ของบริษัทฯ ที่หมายเลข 02-949-1999 (วันจันทร์ถึงศุกร์
                                        ระหว่างเวลา 8.30 – 17.30 น.
                                        ยกเว้นวันหยุดทำการตามประกาศของตลาดหลักทรัพย์แห่งประเทศไทย) </span>
                                </div>
                                <div> <span style='font-size:14.0pt'> <span style='font-family:cordia new,sans-serif'>
                                            &nbsp; </span> </span> </div>
                                <div><span style='font-size:14.0pt'><span
                                            style='font-family:cordia new,sans-serif'>ขอแสดงความนับถือ</span></span>
                                </div>
                                <div><span style='font-size:14.0pt'><span
                                            style='font-family:cordia new,sans-serif'>บริษัทหลักทรัพย์ ไทยพาณิชย์
                                            จำกัด</span></span></div>
                                <div> <span style='font-size:14.0pt'> <span style='font-family:cordia new,sans-serif'>
                                            &nbsp; </span> </span> </div>
                                <div><span style='font-size:14.0pt'><span
                                            style='font-family:cordia new,sans-serif'>To&nbsp;{data.DeliveryInfo.CustomerFullName}</span></span>
                                </div>
                                <div> <span style='font-size:14.0pt'> <span style='font-family:cordia new,sans-serif'>
                                            SCB Securities Company Limited (“SCBS”) has attached your Confirmation
                                            Note/Combined Statement/Payment/Tax Invoice and/or Deposit Slip and/or
                                            Withdrawal Slip and/or Exchange Slip in PDF file format. You can open the
                                            file by using your password which is your last 6 digits of citizen number
                                            (for example, if your citizen number is 1-1115-00240-43-6, your password
                                            will be 240436). <br /> The attached documents are PDF files that can be
                                            viewed with Adobe Reader version 6.0 and up. If you cannot open the files,
                                            please download free program from this link
                                            http://www.adobe.com/products/acrobat/ <br /> This e-mail is distributed
                                            automatically, please do not reply. Any questions or further information,
                                            please feel free to contact Call Center at 02-949-1999 (Monday to Friday
                                            from 8:30 to 17:30, except holidays observed by the Stock Exchange of
                                            Thailand). </span> </span> </div>
                                <div><span style='font-size:14.0pt'><span
                                            style='font-family:cordia new,sans-serif'>Yours sincerely,</span></span>
                                </div>
                                <div><span style='font-size:14.0pt'><span style='font-family:cordia new,sans-serif'>SCB
                                            Securities Company Limited</span></span></div>
                                <div> <span style='font-size:14.0pt'> <span style='font-family:cordia new,sans-serif'>
                                            &nbsp; </span> </span> </div>
                                <div>&nbsp;</div>
                            </td>
                        </tr>
                    </tbody>
                </table>
                <div>&nbsp;</div>
                <table>
                    <tbody>
                        <tr>
                            <td bgcolor='#ffffff'>
                                <font color='#000000'>
                                    <pre>DISCLAIMER: This e-mail is intended solely for the recipient(s) name above. If you are not the intended recipient, any type of your use is prohibited. Any information, comment or statement contained in this e-mail, including any attachments (if any) are those of the author and are not necessarily endorsed by the Bank. The Bank shall, therefore, not be liable or responsible for any of such contents, including damages resulting from any virus transmitted by this e-mail.</pre>
                                </font>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </blockquote>
    </div>
    <div dir='auto'>
        <blockquote type='cite'>
            <div dir='ltr'></div>
        </blockquote>
    </div>
</body>

</html>";
            return template;
        }

        private static string SDC(DocumentMeta data)
        {
            var template = $@"<html> <head> </head> <body> <div dir='auto'> <blockquote type='cite'> <div dir='ltr'> <meta content='text/html; charset=utf-8'> <table border='0' cellpadding='0' cellspacing='0'> <tbody> <tr> <td> <div><span style='font-size:14.0pt'><span style='font-family:cordia new,sans-serif'>เรียน&nbsp;{data.DeliveryInfo.CustomerFullName}</span></span></div> <div> <span style='font-size:14.0pt'> <span style='font-family:cordia new,sans-serif'> ใบยืนยันการซื้อ/ขายหลักทรัพย์/ใบเสร็จรับเงิน/ใบสำคัญจ่ายเงิน/ใบกำกับภาษี และหรือใบนำฝาก/ใบรับเงิน และหรือ ใบคำขอถอนเงิน ที่ท่านได้รับจาก<br> บริษัท หลักทรัพย์ไทยพาณิชย์ จำกัด ตามเอกสารที่แนบมานี้ ท่านสามารถเปิดดูได้ด้วยรหัสผ่านของท่าน ซึ่งรหัสผ่านของท่านคือ เลขวันเดือนปี (ค.ศ.) เกิดของท่าน<br> ในรูปของ DDMMYYYY (เช่น วันเกิดของท่านคือวันที่ 31 มกราคม ค.ศ.1999 รหัสผ่านคือ 31011999) </span> </span> </div> <div> <span style='font-size:14.0pt'> <span style='font-family:cordia new,sans-serif'> &nbsp; </span> </span> </div> <div><span style='font-size:14.0pt'><span style='font-family:cordia new,sans-serif'>ขอแสดงความนับถือ</span></span></div> <div><span style='font-size:14.0pt'><span style='font-family:cordia new,sans-serif'>บริษัทหลักทรัพย์ ไทยพาณิชย์ จำกัด</span></span></div> <div> <span style='font-size:14.0pt'> <span style='font-family:cordia new,sans-serif'> &nbsp; </span> </span> </div> <div><span style='font-size:14.0pt'><span style='font-family:cordia new,sans-serif'>ข้อสังเกต:</span></span></div> <div> <span style='font-size:14.0pt'> <span style='font-family:cordia new,sans-serif'> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 1. อีเมลฉบับนี้เป็นการแจ้งข้อมูลจากระบบโดยอัตโนมัติ กรุณาอย่าตอบกลับ หากท่านมีข้อสงสัยหรือต้องการสอบถามรายละเอียดเพิ่มเติม กรุณาติดต่อ ฝ่ายชำระราคา<br> บริษัทหลักทรัพย์ ไทยพาณิชย์ จำกัด หมายเลขโทรศัพท์ 02-949-1209, 02-949-1205 หรือ ฝ่ายลูกค้าสัมพันธ์ หมายเลขโทรศัพท์ 02-949-1999 (วันจันทร์ถึงศุกร์<br> ระหว่างเวลา 8.30 – 17.30 น. ยกเว้นวันหยุดของตลาดหลักทรัพย์แห่งประเทศไทย) </span> </span> </div> <div> <span style='font-size:14.0pt'> <span style='font-family:cordia new,sans-serif'> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 2. เอกสารที่แนบมานี้จะอยู่ในรูปแบบ PDF ซึ่งเปิดดูได้ด้วยโปรแกรม Adobe Reader เวอร์ชั่น 6.0 ขึ้นไป (ท่านสามารถดาวน์โหลดโปรแกรมได้ที่เว็บไซต์ โดยพิมพ์ URL <a href='http://www.adobe.com' target='_blank'>http://www.adobe.com</a>) </span> </span> </div> <div> <span style='font-size:14.0pt'> <span style='font-family:cordia new,sans-serif'> &nbsp; </span> </span> </div> <div><span style='font-size:14.0pt'><span style='font-family:cordia new,sans-serif'>To&nbsp;{data.DeliveryInfo.CustomerFullName}</span></span></div> <div> <span style='font-size:14.0pt'> <span style='font-family:cordia new,sans-serif'> Confirmation Note/Receipt/ Payment Voucher/Tax Invoice and or Deposit/Receive Slip and or Cash Withdrawal Slip you receive from SCB Securities Company<br> Limited as attached herewith may be viewed with your password which is your birthdate DDMMYYYY (for example, if your birthdate is 31 January 1999, your<br> password will be 31011999). </span> </span> </div> <div><span style='font-size:14.0pt'><span style='font-family:cordia new,sans-serif'>Yours sincerely,</span></span></div> <div><span style='font-size:14.0pt'><span style='font-family:cordia new,sans-serif'>SCB Securities Company Limited</span></span></div> <div> <span style='font-size:14.0pt'> <span style='font-family:cordia new,sans-serif'> &nbsp; </span> </span> </div> <div><span style='font-size:14.0pt'><span style='font-family:cordia new,sans-serif'>Remarks:</span></span></div> <div> <span style='font-size:14.0pt'> <span style='font-family:cordia new,sans-serif'> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 1. This e-mail is distributed automatically, please do not reply. Should you have any questions or require further information, please contact Settlement<br> Department, SCB Securities Company Limited, Tel: 02-949-1209, 02-949-1205 or Call Center, Tel: 02-949-1999 (Monday to Friday from 8:30 to 17:30,<br> except&nbsp; </span> </span><span style='font-size:14.0pt'><span style='font-family:cordia new,sans-serif'>holidays observed by the Stock Exchange of Thailand).</span></span> </div> <div> <span style='font-size:14.0pt'> <span style='font-family:cordia new,sans-serif'> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 2. The attached documents are PDF files that can be viewed with Adobe Reader version 6.0 or later (downloadable at URL <a href='http://www.adobe.com' target='_blank'>http://www.adobe.com</a>). </span> </span> </div> <div>&nbsp;</div> </td> </tr> </tbody> </table> <div>&nbsp;</div> <table> <tbody> <tr> <td bgcolor='#ffffff'> <font color='#000000'> <pre>DISCLAIMER: This e-mail is intended solely for the recipient(s) name above. If you are not the intended recipient, any type of your use is prohibited. Any information, comment or statement contained in this e-mail, including any attachments (if any) are those of the author and are not necessarily endorsed by the Bank. The Bank shall, therefore, not be liable or responsible for any of such contents, including damages resulting from any virus transmitted by this e-mail.</pre> </font> </td> </tr> </tbody> </table> </div> </blockquote> </div> <div dir='auto'> <blockquote type='cite'> <div dir='ltr'></div> </blockquote> </div> </body> </html>";
            return template; 
        }

        private static async Task LogToDB(string json, SendMail data, DateTime date, string msgId, string error, string remark)
        {
            try
            {
                var sql = $@"INSERT INTO 
                                          public.job_step_execution
                                        (
                                          job_execution_id,
                                          data_date,
                                          account,
                                          product,
                                          document_type,
                                          step_name,
                                          status,
                                          error_message,
                                          message_id,
                                          message_payload,
                                          remark,
                                          create_by,
                                          create_datetime
                                        )
                                        VALUES (
                                          '{data.DocumentMeta.JobExecutionId}',
                                          '{date.ToString("MM-dd-yyyy")}',
                                          '{data.DocumentMeta.AccountNumber}',
                                          '{data.DocumentMeta.Product}',
                                          '{data.DocumentMeta.DocumentType}',
                                          'SENDING_EMAIL',
                                          '{(error == null ? $"SUCCESS" : "ERROR")}',
                                          {(error != null ? $"'{error}'" : "null")},
                                          '{msgId}',
                                          '{json}',
                                          {(remark != null ? $"'{remark}'" : "null")},
                                          'Lambda_SendingEmail',
                                          now()
                                        );";

                using (var conn = await Utility.CreateConnection())
                {
                     await conn.ExecuteNonQueryAsync(sql);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
            
        }

    }

    public class SendMail
    {
        public DocumentMeta DocumentMeta { get; set; }
        public Document[] Documents { get; set; }
    }

    public class DocumentMeta
    {
        public int JobExecutionId { get; set; }
        public string DataDate { get; set; }
        public string AccountNumber { get; set; }
        public string Product { get; set; }
        public string DocumentType { get; set; }
        public string ExecutionCutoffTime { get; set; }

        //public string DocumentType { get; set; }
        //public string PrintDelivery { get; set; }
        //public bool SendEmail { get; set; }

        public DeliveryInfo DeliveryInfo { get; set; }
        //public MarketingInfo MarketingInfo { get; set; }
    }

    public class DeliveryInfo
    {
        public string CustomerFullName { get; set; }
        public string CustomerEmail { get; set; }
        public string AccountNumberFormat { get; set; }
    }

    //public class MarketingInfo
    //{
    //    public string MarketingId { get; set; }
    //    public string MarketingName { get; set; }
    //    public string TeamId { get; set; }
    //    public string TeamName { get; set; }
    //    public string Email { get; set; }
    //}

    public class Document
    {
        public string BucketName { get; set; }
        public string Path { get; set; }
    }
}
