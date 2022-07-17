using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Azure.Identity;
using Azure.Storage.Blobs;
using Microsoft.Extensions.Logging;
using PnP.Core.Admin.Model.SharePoint;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Web;
using CsvHelper;
using CsvHelper.Configuration;

namespace Mcd79.M365DataRepo.Functions
{
    public class ListSites
    {
        private readonly ILogger logger;
        private readonly IPnPContextFactory contextFactory;
        private readonly AzureFunctionSettings azureFunctionSettings;

        public ListSites(IPnPContextFactory pnpContextFactory, ILoggerFactory loggerFactory, AzureFunctionSettings settings)
        {
            logger = loggerFactory.CreateLogger<ListSites>();
            contextFactory = pnpContextFactory;
            azureFunctionSettings = settings;
        }

        /// <summary>
        /// Demo function that creates a site collection, uploads an image to site assets and creates a page with an image web part
        /// GET/POST url: http://localhost:7071/api/ListSites
        /// </summary>
        /// <param name="req"></param>
        /// <returns></returns>
        [Function("ListSites")]
        public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "get")] HttpRequestData req)
        {
            logger.LogInformation("ListSites function starting...");

            HttpResponseData response = null;

            try
            {
                using (var pnpContext = await contextFactory.CreateAsync("Default"))
                {
                    response = req.CreateResponse(HttpStatusCode.OK);
                    response.Headers.Add("Content-Type", "application/json");

                   logger.LogInformation($"Loading all sites");

                    // Create the new site collection
                    var siteCollections = await pnpContext.GetSiteCollectionManager().GetSiteCollectionsWithDetailsAsync();

                    //TODO: Not hard code this
                    string accountName = "m365datarepo";
                    string containerName = "m365datarepo";
                    string containerEndpoint = string.Format("https://{0}.blob.core.windows.net/{1}",
                                                accountName,
                                                containerName);
/*
                    foreach(var siteCollection in siteCollections) {
                        // Get a credential and create a service client object for the blob container.
                        BlobContainerClient containerClient = new BlobContainerClient(new Uri(containerEndpoint),
                                                                    new DefaultAzureCredential());

                        try
                        {
                            // Create the container if it does not exist.
                            await containerClient.CreateIfNotExistsAsync();
                            string blobName = $"sites/{siteCollection.Id.ToString()}.csv";

                            // Upload text to a new block blob.
                            string blobContents = JsonSerializer.Serialize(siteCollection);
                            byte[] byteArray = Encoding.ASCII.GetBytes(blobContents);

                            using (MemoryStream stream = new MemoryStream(byteArray))
                            {
                                await containerClient.UploadBlobAsync(blobName, stream);
                            }
                        }
                        catch (Azure.RequestFailedException e)
                        {
                            Console.WriteLine(e.Message);
                            Console.ReadLine();
                            throw;
                        }
                    }
*/
                    using (var memoryStream = new MemoryStream())
                            {
                                using (var streamWriter = new StreamWriter(memoryStream))
                                using (var csvWriter = new CsvWriter(streamWriter))
                                {
                                    csvWriter.WriteRecords(siteCollections);
                                } // StreamWriter gets flushed here.
                                string csvText = System.Text.Encoding.GetString(memoryStream.ToArray());
                                await response.WriteStringAsync(csvText);

                                return memoryStream.ToArray();
                            }
                    //logger.LogInformation($"Sites created: {siteCollections.Count} found");


                    // Return the URL of the created site
                    

                    return response;
                }
            }
            catch (Exception ex)
            {
                response = req.CreateResponse(HttpStatusCode.OK);
                response.Headers.Add("Content-Type", "application/json");
                await response.WriteStringAsync(JsonSerializer.Serialize(new { error = ex.Message }));
                return response;
            }
        }
    }
}