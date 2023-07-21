
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace WebCustomization.ExcelUpdate
{
    public static class ExcelUpdate
    {
        [FunctionName("ExcelUpdate")]
        public static async Task<HttpResponseMessage> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            string requestBody = String.Empty;

            log.LogInformation("ExcelUpdate - Request body received.");
            // Populate request body
            using (StreamReader streamReader = new StreamReader(req.Body))
            {
                requestBody = await streamReader.ReadToEndAsync();
            }

            log.LogInformation("ExcelUpdate - JSON deserialised and parsed through model.");
            List<Root> myDeserializedClass = JsonConvert.DeserializeObject<List<Root>>(requestBody);

            log.LogInformation("ExcelUpdate - Generating file");
            MemoryStream memoryStream = Helper.generateExcel(myDeserializedClass);

            //Create the response to return
            HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);

            //Set the PDF document content response
            response.Content = new ByteArrayContent(memoryStream.ToArray());

            //Set the contentDisposition as attachment
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = "Output.xlsx"
            };
            //Set the content type as PDF format mime type
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            log.LogInformation("ExcelUpdate - Sending response to the request");
            //Return the response with output excel stream
            return response;
        }
    }
}
