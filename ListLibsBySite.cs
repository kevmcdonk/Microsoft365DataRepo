using System;
using System.IO;
using CsvHelper;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace Mcd79.M365DataRepo.Functions
{
    public class ListLibsBySite
    {
        private readonly ILogger _logger;

        public ListLibsBySite(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<ListLibsBySite>();
        }

        [Function("ListLibsBySite")]
        public void Run([BlobTrigger("samples-workitems/{name}", Connection = "m365datarepo_STORAGE")] string myBlob, string name)
        {
            logger.LogInformation($"ListLIbrariesBySite function starting for {siteBlob}...");

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

                    foreach(var list in pnpContext.Web.Lists.AsRequested())
                    {
                                                // Get a credential and create a service client object for the blob container.
                        BlobContainerClient containerClient = new BlobContainerClient(new Uri(containerEndpoint),
                                                                    new DefaultAzureCredential());

                        try
                        {
                            // Create the container if it does not exist.
                            await containerClient.CreateIfNotExistsAsync();
                            string blobName = $"lists/{list.Id.ToString()}.json";

                            // Upload text to a new block blob.
                            string blobContents = JsonSerializer.Serialize(list);
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


                    // Return the URL of the created site
                    await response.WriteStringAsync("Lists captured");

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
