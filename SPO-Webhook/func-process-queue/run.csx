using Microsoft.ServiceBus.Messaging;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

public static void Run(string myQueueItem, TraceWriter log)
{
    log.Info($"C# ServiceBus queue trigger function processed message: {myQueueItem}");


    //Get values from Application Settings
    string aadAppId = System.Environment.GetEnvironmentVariable("SPO_ApplicationId"); // ID from App registration in Azure AD
    string certThumbprint = System.Environment.GetEnvironmentVariable("SPO_Cert_Thumbprint"); // Certificate Thumbprint
    string certName = System.Environment.GetEnvironmentVariable("SPO_Cert_Name"); // PFX file name:  az-func-webhook.pfx
    string certPassword = System.Environment.GetEnvironmentVariable("SPO_Cert_Password");  // PFX Password 
    string tenant = System.Environment.GetEnvironmentVariable("SPO_Tenant"); //ex: mytenant.onmicrosoft.com
    string siteUrl = System.Environment.GetEnvironmentVariable("SPO_SiteUrl"); // SPO Site collection URL : https://lme.sharepoint.com
    string csEndpoint = System.Environment.GetEnvironmentVariable("CS_Endpoint"); //ex: https://westeurope.api.cognitive.microsoft.com/
    string csAccessKey = System.Environment.GetEnvironmentVariable("CS_AccessKey"); // The access key I get from Cognitive Services


    X509Certificate2 cert = new X509Certificate2(Path.Combine(exCtx.FunctionDirectory, certName), certPassword);

    //Get webhook content
    ResponseModel<NotificationModel> notifications = JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(myQueueItem);

    //In case there's more than one notification...
    foreach (var notification in notifications.Value)
    {
        string listId = notification.Resource;
        string webId = notification.WebId;

        using (ClientContext cc = new AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(siteUrl, aadAppId, tenant, cert))
        {
            //Get related list
            Web web = cc.Site.OpenWebById(new Guid(webId));
            List list = web.Lists.GetById(new Guid(listId));
            cc.Load(list);
            cc.ExecuteQueryRetry();

            //Query the changelog of the list
            //The method requires a Start Token, which is the list.CurrentChangeToken saved in the property bag of the list the last time the function ran
            ChangeToken currentChangeToken = list.CurrentChangeToken;
            ChangeToken startToken = new ChangeToken();
            ChangeQuery query = new ChangeQuery(false, false);

            string token = list.GetPropertyBagValueString("lastChangeToken", "");
            if (!string.IsNullOrEmpty(token))
            {
                startToken.StringValue = token;
                query.ChangeTokenStart = startToken;
            }

            // Limit result to Item Added
            query.Add = true;
            query.Item = true;

            var changesColl = list.GetChanges(query);
            cc.Load(changesColl);
            cc.ExecuteQueryRetry();

            //Process changes
            foreach (Change change in changesColl)
            {
                if (change is ChangeItem)
                {
                    try
                    {

                        // Get List Item
                        ChangeItem item = (ChangeItem)change;
                        ListItem listItem = list.GetItemById(item.ItemId);
                        cc.Load(listItem);
                        cc.ExecuteQueryRetry();

                        // Get the file
                        Microsoft.SharePoint.Client.File file = listItem.File;
                        ClientResult<Stream> fileStreamResult = file.OpenBinaryStream();
                        cc.Load(file);
                        cc.ExecuteQueryRetry();


                        //Convert file to Byte Array Content, to send it to Bing Cognitives Services
                        Stream fileStream = fileStreamResult.Value;
                        BinaryReader binaryReader = new BinaryReader(fileStream);
                        byte[] byteData = binaryReader.ReadBytes((int)fileStream.Length);

                        //Process image using Bing Cognitives Services
                        RestClient client = new RestClient(csEndpoint);
                        RestRequest request = new RestRequest("/vision/v1.0/analyze?visualFeatures=Categories,Description,Tags&language=en");
                        request.AddHeader("Ocp-Apim-Subscription-Key", csAccessKey);
                        request.AddFile("files", byteData, file.Title, "application/octet-stream");
                        var response = client.Post(request);
                        dynamic responseContent = JsonConvert.DeserializeObject(response.Content);
                        string description = responseContent.description.captions[0].text.Value;
                        log.Info("Description : " + description);

                        //Writeback description to the list item in SharePoint
                        listItem["Description"] = description;
                        listItem.Update();
                        cc.ExecuteQueryRetry();
                    }
                    catch(Exception ex)
                    {
                        log.Info(ex.Message);
                    }
                }
            }
            if (changesColl.Count() > 0)
            {
                list.SetPropertyBagValue("lastChangeToken", currentChangeToken.StringValue);
            }
        }
    }
}

public class ResponseModel<T>
{
    [JsonProperty(PropertyName = "value")]
    public List<T> Value { get; set; }
}

public class NotificationModel
{
    [JsonProperty(PropertyName = "subscriptionId")]
    public string SubscriptionId { get; set; }

    [JsonProperty(PropertyName = "clientState")]
    public string ClientState { get; set; }

    [JsonProperty(PropertyName = "expirationDateTime")]
    public DateTime ExpirationDateTime { get; set; }

    [JsonProperty(PropertyName = "resource")]
    public string Resource { get; set; }

    [JsonProperty(PropertyName = "tenantId")]
    public string TenantId { get; set; }

    [JsonProperty(PropertyName = "siteUrl")]
    public string SiteUrl { get; set; }

    [JsonProperty(PropertyName = "webId")]
    public string WebId { get; set; }
}

public class SubscriptionModel
{
    [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
    public string Id { get; set; }

    [JsonProperty(PropertyName = "clientState", NullValueHandling = NullValueHandling.Ignore)]
    public string ClientState { get; set; }

    [JsonProperty(PropertyName = "expirationDateTime")]
    public DateTime ExpirationDateTime { get; set; }

    [JsonProperty(PropertyName = "notificationUrl")]
    public string NotificationUrl { get; set; }

    [JsonProperty(PropertyName = "resource", NullValueHandling = NullValueHandling.Ignore)]
    public string Resource { get; set; }
}
