using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

using Amazon.Lambda.Core;
using Amazon.Lambda.SNSEvents;
using System.Text.Json;
using System.Data.Common;
using System.Runtime.CompilerServices;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace EmailNotify
{
    public class Function
    {

        static int cutoff = int.Parse(Utility.Env("CUTOFF_TIME"));
        static DateTime now = DateTime.MinValue;

        /// <summary>
        /// A simple function that takes a string and does a ToUpper
        /// </summary>
        /// <param name="input"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        public async Task FunctionHandler(SNSEvent snsEvent, ILambdaContext context)
        {
            now = DateTime.Now.AddHours(7.0);
            foreach (var item in snsEvent.Records)
            {
                await Process(item.Sns.Message);
            }
        }

        public async Task Process(string json)
        {

            var data = JsonDocument.Parse(json).RootElement;
            var msgId = data.GetProperty("mail").GetProperty("messageId").GetString();



            var check = await CheckDuplicate(msgId);
            if (!check)
            {
                var header = headerExtract(data);
                header.MessageId = msgId;
                if (!header.ExecutionCutoffTime.HasValue) {
                    header.ExecutionCutoffTime = header.DataDate.AddMinutes(cutoff);
                }

                var state = data.GetProperty("notificationType").GetString();

                switch (state)
                {
                    case "Delivery":
                        await Delivery(json, header);
                        break;
                    case "Bounce":
                        await Bounce(json, data, header);
                        break;
                }
            }
            

        }

        private static async ValueTask<bool> CheckDuplicate(string msgId)
        {
            var sql = $@"SELECT 1 FROM public.job_step_execution WHERE message_id = '{msgId}' LIMIT 1";
            using (var conn = await Utility.CreateConnection())
            {
                var result = await conn.ExecuteScalarAsync(sql);
                return result != null && (int)result != 0;
            }
        }

        private async Task Delivery(string json,Header header)
        {
            
            if (header.ExecutionCutoffTime >= now)
            {
                await LogToDB(json, header, "SUCCESS", null);
            }
            else
            {
                await LogToDB(json, header, "CO_SUCCESS", $"cutoff: {header.ExecutionCutoffTime.Value.ToString("dd/MM/yyyy HH:mm:ss")} current: {now.ToString("dd/MM/yyyy HH:mm:ss")}");
            }

            //await LogToDB(json, header, "SUCCESS", null);
        }

        private async Task Bounce(string json, JsonElement data, Header header)
        {
            var bounce = data.GetProperty("bounce");
            var type = bounce.GetProperty("bounceType").GetString();
            var subType = bounce.GetProperty("bounceSubType").GetString();
            

            if (header.ExecutionCutoffTime >= now)
            {
                await LogToDB(json, header, "BOUNCE", $"{type}:{subType}");
            }
            else
            {
                await LogToDB(json, header, "CO_BOUNCE", $"{type}:{subType}");
            }
        }

        private Header headerExtract(JsonElement data)
        {
            var header = new Header();
            foreach (var item in data.GetProperty("mail").GetProperty("headers").EnumerateArray())
            {
                var name = item.GetProperty("name").GetString();
                if (name == "DataDate")
                {
                    header.DataDate = DateTime.ParseExact(item.GetProperty("value").GetString(), "dd/MM/yyyy", null);
                }
                else if (name == "Product")
                {
                    header.Product = item.GetProperty("value").GetString();
                }
                else if (name == "DocumentType")
                {
                    header.DocumentType = item.GetProperty("value").GetString();
                }
                else if (name == "JobExecutionId")
                {
                    header.JobExecutionId = int.Parse(item.GetProperty("value").GetString());
                }
                else if (name == "AccountNumber")
                {
                    header.AccountNumber = item.GetProperty("value").GetString();
                }
                else if (name == "ExecutionCutoffTime")
                {
                    try {
                        header.ExecutionCutoffTime = DateTime.ParseExact(item.GetProperty("value").GetString(), "yyyy-MM-dd HH:mm", null);
                    }
                    catch (FormatException) {
                        header.ExecutionCutoffTime = null;
                    }
                }
            }
            return header;
        }

        private static async Task LogToDB(string json, Header header, string state, string remark)
        {
            try
            {
                var sql= $@"INSERT INTO 
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
                                          '{header.JobExecutionId}',
                                          '{header.DataDate.ToString("MM-dd-yyyy")}',
                                          '{header.AccountNumber}',
                                          '{header.Product}',
                                          '{header.DocumentType}',
                                          'SENT_EMAIL',
                                          '{state}',
                                          null,
                                          '{header.MessageId}',
                                          '{json}',
                                          {(remark != null ? $"'{remark}'" : "null")},
                                          'LAMBDA_SENT_EMAIL',
                                          now()
                                        );";
                using (var conn = await Utility.CreateConnection())
                {            
                    await conn.ExecuteNonQueryAsync(sql);
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }

        }

       

    }

    public class Header
    {
        public DateTime DataDate { get; set; }
        public string Product { get; set; }
        public string DocumentType { get; set; }
        public int JobExecutionId { get; set; }
        public string MessageId { get; set; }
        public string AccountNumber { get; set; }
        public DateTime? ExecutionCutoffTime { get; set; }
    }
  
}
