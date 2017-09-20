using System.Net;
using Microsoft.ServiceBus.Messaging;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info("Incoming notification");

    // Get Validation Token
    string validationToken = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "validationtoken", true) == 0)
        .Value;

    // Response to SharePoint including the validation token
    if (validationToken != null)
    {
        log.Info($"Validation token {validationToken} received");
        var response = req.CreateResponse(HttpStatusCode.OK);
        response.Content = new StringContent(validationToken);
        return response;
    }

    var content = await req.Content.ReadAsStringAsync();
    log.Info($"Received following payload: {content}");

    var connStr = System.Environment.GetEnvironmentVariable("connStr_WebHookQueue");
    var client = QueueClient.CreateFromConnectionString(connStr);
    var message = new BrokeredMessage(content);
    client.Send(message);

    return new HttpResponseMessage(HttpStatusCode.OK);
}
