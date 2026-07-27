using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;

namespace Zotero_linker_ppt
{
    internal sealed class ZoteroHttpIntegrationClient
    {
        private const string ExecCommandUrl = "http://127.0.0.1:23119/connector/document/execCommand";
        private const string RespondUrl = "http://127.0.0.1:23119/connector/document/respond";
        private const int MaxSteps = 200;

        private readonly JavaScriptSerializer serializer;

        internal ZoteroHttpIntegrationClient()
        {
            serializer = new JavaScriptSerializer
            {
                MaxJsonLength = int.MaxValue
            };
        }

        internal void Execute(string command, PowerPointZoteroDocument document)
        {
            if (document == null)
            {
                throw new ArgumentNullException("document");
            }

            document.CurrentIntegrationCommand = command;
            string response = PostJson(ExecCommandUrl, new Dictionary<string, object>
            {
                { "command", command },
                { "docId", document.DocumentId }
            });

            for (int step = 0; step < MaxSteps; step += 1)
            {
                ZoteroIntegrationRequest request = ParseRequest(response);
                if (request == null)
                {
                    return;
                }

                object result;
                try
                {
                    result = document.Execute(request.Command, request.Arguments);
                }
                catch (Exception ex)
                {
                    result = new Dictionary<string, object>
                    {
                        { "error", "Connector Error" },
                        { "message", ex.Message },
                        { "stack", ex.ToString() }
                    };
                }

                if (IsCompleteCommand(request.Command))
                {
                    return;
                }

                response = PostRawJson(RespondUrl, SerializeResult(result));
            }

            throw new InvalidOperationException("Zotero integration transaction did not finish.");
        }

        private ZoteroIntegrationRequest ParseRequest(string response)
        {
            if (string.IsNullOrWhiteSpace(response))
            {
                return null;
            }

            object parsed = serializer.DeserializeObject(response);
            string inner = parsed as string;
            if (inner != null)
            {
                parsed = serializer.DeserializeObject(inner);
            }

            Dictionary<string, object> value = parsed as Dictionary<string, object>;
            if (value == null || !value.ContainsKey("command"))
            {
                return null;
            }

            return new ZoteroIntegrationRequest
            {
                Command = Convert.ToString(value["command"], System.Globalization.CultureInfo.InvariantCulture),
                Arguments = value.ContainsKey("arguments") ? value["arguments"] as object[] : new object[0]
            };
        }

        private string SerializeResult(object result)
        {
            return result == null ? "null" : serializer.Serialize(result);
        }

        private static bool IsCompleteCommand(string command)
        {
            return string.Equals(command, "Document.complete", StringComparison.OrdinalIgnoreCase);
        }

        private string PostJson(string url, object payload)
        {
            return PostRawJson(url, serializer.Serialize(payload));
        }

        private static string PostRawJson(string url, string json)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(json ?? "null");
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Accept = "application/json";
            request.ContentLength = bytes.Length;

            try
            {
                using (Stream stream = request.GetRequestStream())
                {
                    stream.Write(bytes, 0, bytes.Length);
                }

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (WebException ex)
            {
                HttpWebResponse response = ex.Response as HttpWebResponse;
                if (response != null && response.StatusCode == HttpStatusCode.ServiceUnavailable)
                {
                    throw new InvalidOperationException("Zotero is already processing another citation command. Restart Zotero if this persists.", ex);
                }

                throw new InvalidOperationException("Could not communicate with Zotero. Please start Zotero and enable local application communication.", ex);
            }
        }
    }

    internal sealed class ZoteroIntegrationRequest
    {
        internal string Command { get; set; }
        internal object[] Arguments { get; set; }
    }
}
