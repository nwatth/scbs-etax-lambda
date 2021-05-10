using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Amazon;
using Amazon.Lambda.Core;
using Amazon.S3;
using Amazon.S3.Model;
using System.Text;
using System.Runtime.CompilerServices;
using System.Text.Json;


// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace EmailStatus
{
    public class Function
    {

        static JsonSerializerOptions serializerOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        };
        public async Task<string> FunctionHandler(string input, ILambdaContext context)
        {
            
            var manual = false;
            var date = DateTime.MinValue;
            if (string.IsNullOrEmpty(input))
            {
                date = await LoadFindMaxDate();
            }
            else
            {
                date = DateTime.ParseExact(input, "dd/MM/yyyy", null);
                manual = true;
            }
            var dateText = date.ToString("yyyyMMdd");
            var datePG = date.ToString("MM-dd-yyyy");
            if (await ExitInS3(dateText) == false || manual)
            {
                var successNum = 0;
                var failNum = 0;
                var successBuilder = new StringBuilder();
                var failBuilder = new StringBuilder();
                var summary = new List<(string date, string doc, int success, int fail)>();
                var data = await LoadStepData(datePG);
                var doc = data.Count != 0 ? data[0].doc : null;
                
                
                for (int i = 0; i < data.Count; i++)
                {
                    var item = data[i];
                    if(doc != item.doc)
                    {
                        summary.Add((datePG, doc, successNum, failNum));
                        doc = item.doc;
                        successNum = 0;
                        failNum = 0;
                    }


                    if (item.status)
                    {
                        successNum++;
                        successBuilder.Append($"{item.account}|{item.doc}\r\n");
                    }
                    else
                    {
                        failNum++;
                        failBuilder.Append($"{item.account}|{item.doc}\r\n");
                        
                    }
                }

                if(data.Count != 0)
                {
                    summary.Add((datePG, doc, successNum, failNum));
                }

                var tasks = new Task[3];

            tasks[0] = WriteToS3(dateText, "success", successBuilder.ToString());
            tasks[1] = WriteToS3(dateText, "fail", failBuilder.ToString());
            tasks[2] = WriteSummaryToDB(summary);
            await Task.WhenAll(tasks);
            }

            return "success";
        }

        public Task WriteToS3(string dateText, string status, string data)
        {
            var client = new AmazonS3Client(RegionEndpoint.APSoutheast1);
            var request = new PutObjectRequest();
            request.BucketName = Environment.GetEnvironmentVariable("BUCKET");
            request.ContentType = "text/plain";
            request.Key = $"{Environment.GetEnvironmentVariable("PATH")}{status}_{dateText}.txt";
            request.ContentBody = data;
            
            return client.PutObjectAsync(request);
        }

        private static List<Task> DeletePdf(JsonElement root)
        {
            var tasks = new List<Task>();
            var doc = JsonSerializer.Deserialize<Document[]>(root.GetProperty("documents").GetRawText(), serializerOptions);
            var s3Client = new AmazonS3Client(RegionEndpoint.APSoutheast1);

            foreach (var item in doc)
            {
                tasks.Add(s3Client.DeleteObjectAsync(item.BucketName, item.Path));
            }
            return tasks;
        }

        private static List<Task> DeleteArchivePdf(JsonElement root)
        {
            var tasks = new List<Task>();
            var s3Client = new AmazonS3Client(RegionEndpoint.APSoutheast1);

            foreach (var item in root.GetProperty("documents").EnumerateArray())
            {
                var doc = item.GetProperty("documentConfig");
                var file = doc.GetProperty("fileName").GetString();

                foreach (var archive in doc.GetProperty("archivePaths").EnumerateArray())
                {
                    var bucket = archive.GetProperty("bucketName").GetString();
                    foreach (var path in archive.GetProperty("path").EnumerateArray())
                    {
                        var p = path.GetString() + file;
                        tasks.Add(s3Client.DeleteObjectAsync(bucket, p));
                    }


                }

            }

            return tasks;


        }

        public async Task WriteSummaryToDB(List<(string date, string doc, int success, int fail)> data)
        {
            var sqlBuilder = new StringBuilder();
            foreach (var item in data)
            {
                sqlBuilder.Append($@"DELETE FROM public.email_status_summary WHERE data_date = '{item.date}' AND document_type = '{item.doc}';
                                    INSERT INTO 
                                      public.email_status_summary
                                    (
                                      data_date,
                                      document_type,
                                      success_num,
                                      fail_num,
                                      create_by,
                                      create_datetime
                                    )
                                    VALUES (
                                      '{item.date}',
                                      '{item.doc}',
                                      {item.success},
                                      {item.fail},
                                      'Lambda_EmailStatus',
                                      now()
                                    ); ");
            }
            if(sqlBuilder.Length != 0)
            {
                using (var conn = await Utility.CreateConnection())
                {
                    await conn.ExecuteNonQueryAsync(sqlBuilder.ToString());
                }
            }
            
        }

        public async Task<bool> ExitInS3(string dateText)
        {
            try
            {
                var client = new AmazonS3Client(RegionEndpoint.APSoutheast1);
                var request = new GetObjectMetadataRequest();
                request.BucketName = Environment.GetEnvironmentVariable("Bucket");
                request.Key = $"{Environment.GetEnvironmentVariable("PATH")}success_{dateText}.txt";
                var result = await client.GetObjectMetadataAsync(request);
                return true;
            }
            catch
            {
                return false;
            }
            
        }

        public async Task<DateTime> LoadFindMaxDate()
        {
           var sql = $@"SELECT 
                                      data_date
                                    FROM 
                                      public.job_step_execution 
                                    WHERE
                                      document_type = 'SDC' AND product = 'EQUITY'  
                                      ORDER BY data_date DESC LIMIT 1;";

            using (var conn = await Utility.CreateConnection())
            {
                Console.WriteLine($"{conn.ConnectionString}");
                return (DateTime)conn.ExecuteScalar(sql);
            }

           

        }

        public async Task<List<(string account, string doc, bool status)>> LoadStepData(string date)
        {
            //var dateText = date.ToString("MM-dd-yyyy");
            var sql = $@"WITH data 
                                             AS (SELECT account,
                                                        document_type,
                                                        step_name,
                                                        status,
                                                        Row_number() over ( 
                                                            PARTITION BY account 
                                                            ORDER BY id DESC) AS num 
                                                 FROM   public.job_step_execution
                                                 WHERE  step_name = 'SENDING_EMAIL' OR step_name = 'SENT_EMAIL' AND document_type = 'SDC' AND product = 'EQUITY' )
                                        SELECT * 
                                        FROM   data 
                                        WHERE  num = 1
                                        ORDER BY document_type";

            var result = new List<(string account, string doc, bool status)>();

            using var conn = await Utility.CreateConnection();
            var accountSB = new StringBuilder();
            var sending = new HashSet<string>();
            var fail = new HashSet<string>();

            using (var reader = conn.ExecuteReader(sql))
            {

                while (reader.Read())
                {
                    var account = reader.GetString(0);
                    var docType = reader.GetString(1);
                    var step = reader.GetString(2);
                    var status = reader.GetString(3);
                    var check = step == "SENT_EMAIL" && status == "SUCCESS" ? true : false;
  
                    if (!check && step == "SENDING_EMAIL" && status == "SUCCESS")
                    {
                        //fail.Add($"{account}{doc}");
                        accountSB.Append($",'{account}'");
                        sending.Add($"{account}{docType}");

                    }

                    result.Add((account, docType, check));
                }

            }


            if (accountSB.Length != 0)
            {
                var sqlBuild = new StringBuilder();
                sql = $@"SELECT account,
                                                        document_type,
                                                        message_payload,
                                                        status
                                                 FROM   public.job_step_execution
                                                 WHERE  step_name = 'SENDING_EMAIL' AND document_type = 'SDC' AND product = 'EQUITY' AND account IN ({accountSB.Remove(0, 1)})";
                using (var reader = conn.ExecuteReader(sql))
                {
                    
                    while (reader.Read())
                    {
                        var account = reader.GetString(0);
                        var docType = reader.GetString(1);
                        if (sending.Contains($"{account}{docType}"))
                        {
                            var json = reader.GetString(2);
                            var data = JsonDocument.Parse(json).RootElement.GetProperty("documentMeta");
                            var jobId = data.GetProperty("jobExecutionId").GetInt32();
                            var product = data.GetProperty("product").GetString();

                            sqlBuild.Append($"INSERT INTO public.job_step_execution(job_execution_id, data_date, account, product, document_type, step_name, status, error_message, message_id, message_payload, remark, create_by, create_datetime) VALUES('{jobId}', '{date}', '{account}', '{product}', '{docType}','STATUS_EMAIL', 'CO', null, null, '{json}', null, 'LAMBDA_STATUS_EMAIL', now());");
                        }
                    }
                }

                if(sqlBuild.Length != 0)
                {
                    conn.ExecuteNonQuery(sqlBuild.ToString());
                }

                //sql = $@"SELECT account,
                //                                        document_type,
                //                                        message_payload
                                                        
                //                                 FROM   public.job_step_execution
                //                                 WHERE  step_name = 'PUSH_PDF_ARCHIVED_QUEUE' AND account IN ({accountSB})";
                //using (var reader = conn.ExecuteReader(sql))
                //{

                //    while (reader.Read())
                //    {
                //        var account = reader.GetString(0);
                //        var doc = reader.GetString(1);
                //        if (fail.Contains($"{account}{doc}"))
                //        {
                //            var data = JsonDocument.Parse(reader.GetString(2)).RootElement;
                //            tasks.AddRange(DeleteArchivePdf(data));
                //        }
                //    }

                //}

                //await Task.WhenAll(tasks);
            }

            return result;
        }
    }

    public class Document
    {
        public string BucketName { get; set; }
        public string Path { get; set; }
    }
}
