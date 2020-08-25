using System.Collections.Generic;
using System.Linq;
using System.Text;
using RestSharp;
using System.Xml.Linq;
using System.Net;

namespace System.Web.Services.Ellipse.Post
{
    public class PostService
    {
        public string Username { get; set; }
        private string Password { get; set; }
        public string Position { get; set; }
        public string District { get; set; }
        public string Url { get; set; }
        public string ConnectionId { get; set; }
        private RestClient Client { get; set; }
        private RestRequest Request { get; set; }

        public PostService(string userName, string userPswd, string userPost, string userDstr, string url)
        {
            this.Username = userName;
            this.Password = userPswd;
            this.Position = userPost;
            this.District = userDstr;
            this.Url = url;
        }

        public ResponseDto InitConexion()
        {
            var responseDto = new ResponseDto();
            try
            {
                var connectionId = GetNewConnectionId();
                var requestXml = new StringBuilder("");
                requestXml.AppendLine("<interaction>");
                requestXml.AppendLine("<actions>");
                requestXml.AppendLine("<action>");
                requestXml.AppendLine("<name>login</name>");
                requestXml.AppendLine("<data>");
                requestXml.Append("<username>");
                requestXml.Append(this.Username);
                requestXml.AppendLine("</username>");
                requestXml.Append("<password>");
                requestXml.Append(this.Password);
                requestXml.AppendLine("</password>");
                requestXml.Append("<scope>");
                requestXml.Append(this.District);
                requestXml.AppendLine("</scope>");
                requestXml.Append("<position>");
                requestXml.Append(this.Position);
                requestXml.AppendLine("</position>");
                requestXml.AppendLine("<rememberMe>N</rememberMe>");
                requestXml.AppendLine("</data>");
                requestXml.Append("<id>");
                requestXml.Append(connectionId);
                requestXml.AppendLine("</id>");
                requestXml.AppendLine("</action>");
                requestXml.AppendLine("</actions>");
                requestXml.AppendLine("<chains />");
                requestXml.AppendLine("<application>login</application>");
                requestXml.AppendLine("<applicationPage />");
                requestXml.AppendLine("</interaction>");
                var requestLogin = requestXml.ToString();
                Client = new RestClient(this.Url);
                Request = new RestRequest(Method.POST);
                Request.AddHeader("content-type", "application/xml");
                Request.AddParameter("application/xml", requestLogin, ParameterType.RequestBody);
                var response = Client.Execute(Request);
                if (response.ResponseStatus.Equals(ResponseStatus.Completed))
                {
                    var xdoc = XDocument.Parse(response.Content);
                    if (xdoc.Root != null)
                    {
                        var elements = xdoc.Root.Descendants(XName.Get("errors"));
                        responseDto.Errors = Message.GetMessagesByXElements(elements);
                        responseDto.ResponseXML = xdoc;
                        responseDto.ResponseString = xdoc.ToString();
                        var connectionIdElements = xdoc.Root.Descendants(XName.Get("connectionId"));
                        ConnectionId = connectionIdElements.First().Value;
                    }
   
                    var cookieSession = new CookieContainer();
                    foreach (var cookie in response.Cookies)
                    {
                        cookieSession.Add(new Cookie(cookie.Name, cookie.Value, cookie.Path, cookie.Domain));
                    }
                    Client.CookieContainer = cookieSession;
                }
                else
                {
                    throw new Exception(response.ErrorMessage);
                }
            }
            catch (Exception e)
            {
                responseDto.Errors = new List<Message>() {
                    new Message("CatchException", "0", e.StackTrace, e.Message)
                };
            }
            return responseDto;
        }

        public ResponseDto ExecutePostRequest(string xmlRequest)
        {
            var responseDto = new ResponseDto();
            try
            {
                Request.AddHeader("content-type", "application/xml");
                Request.Parameters.RemoveAll(parameter => parameter.Type.Equals(ParameterType.RequestBody));
                Request.AddParameter("application/xml", xmlRequest, ParameterType.RequestBody);
                var response = Client.Execute(Request);
                if (response.ResponseStatus.Equals(ResponseStatus.Completed))
                {
                    var xdoc = XDocument.Parse(response.Content);
                    if (xdoc.Root != null)
                    {
                        var elements = xdoc.Root.Descendants(XName.Get("errors"));
                        responseDto.Errors = Message.GetMessagesByXElements(elements);
                        var informationElements = xdoc.Root.Descendants(XName.Get("information"));
                        responseDto.Informations = Message.GetMessagesByXElements(informationElements);
                    }

                    responseDto.ResponseXML = xdoc;
                    responseDto.ResponseString = xdoc.ToString();
                }
                else
                {
                    throw new Exception(response.ErrorMessage);
                }
            }
            catch (Exception e)
            {
                responseDto.Errors = new List<Message>() {
                    new Message("CatchException", "0", e.StackTrace, e.Message)
                };
            }
            return responseDto;
        }
        public static string GetNewConnectionId()
        {
            return Guid.NewGuid().ToString().ToUpper();
        }
    }
}
