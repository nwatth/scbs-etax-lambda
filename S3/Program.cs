using Amazon;
using Amazon.S3;
using Amazon.S3.Model;
using System;
using System.Threading.Tasks;

namespace S3
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                var awsCredentials = new Amazon.Runtime.BasicAWSCredentials("AKIAW22ZAFMF33PMHEMU", "hBr7/VVdtKD6d2yMsz8tHLvv48QfjmD9FexpowKN");
                var client = new AmazonS3Client(awsCredentials,RegionEndpoint.APSoutheast1);
                //var client = new AmazonS3Client(RegionEndpoint.APSoutheast1);
                var request = new GetObjectMetadataRequest();
                request.BucketName = "s86555";
                request.Key = $"scbcorp/mail-result/fail_20201222.txt";
                var result = await client.GetObjectMetadataAsync(request);
                Console.WriteLine(result.ContentLength);
                Console.ReadLine();
            }
            catch(Exception ex)
            {
                var dd = ex;
            }
            
        }
    }
}
