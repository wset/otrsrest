﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Collections.Specialized;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO;

namespace otrsrest
{
    // OTRS REST Connector.
    //
    // Create new tickets using the TicketCreator class
    //
    // TicketCreator ticket = new TicketCreator();
    //
    //
    // Setup details using the request propert
    //
    // ticket.request.Ticket.Title = ...
    // ticket.request.Ticket.CustomerUser = ... (optional defaults to "oc")
    // ticket.request.Ticket.Queue = ... (optional defaults to "Online Classroom")
    // ticket.request.Ticket.State = ... (optional defaults to "new")
    // ticket.request.Ticket.PriorityID = ... (optional defaults to 3)
    //
    // ticket.request.Article.Subject = ...
    // ticket.request.Article.Body = ...
    // ticket.request.Article.SenderType = ... (optional defaults to "system")
    // ticket.request.Article.ContentType = ... (optional defaults to "text/plain; charset=utf8")
    //
    // Dynamic fields can be added as
    // ticket.request.AddDynamicField(string name, string value)
    //
    // Attachments can be added as
    // ticket.request.AddAttachment(string path)
    //
    // Use async Request() method to make the request (this returns the http status code)
    //
    // await ticket.Request();
    //
    // 
    // The rest response can be access using the response property
    //
    // ticket.response.TicketNumber
    // ticket.response.TicketID
    // ticket.response.ArticleID


    public class NewTicket
    {
        // Class to hold new ticket information.
        // Includes Title, CustomerUser, Queue and State as string properties.
        // and PriorityID as int
        public string Title { get; set; }
        public string CustomerUser { get; set; }
        public string Queue { get; set; }
        public string State { get; set; }
        public int PriorityID { get; set; }

        public NewTicket()
        {
            // Set default values for CustomerUser, Queue, State and PriorityID
            CustomerUser = "oc";
            Queue = "Online Classroom";
            State = "new";
            PriorityID = 3;
        }
    }

    public class NewArticle
    {
        // Class to hold new article information.
        // Includes Subject, Body, SenderType and ContentType as string properties.
        public string Subject { get; set; }
        public string Body { get; set; }
        public string SenderType { get; set; }
        public string ContentType { get; set; }

        public NewArticle()
        {
            // Set default values for SenderType and ContentType.
            SenderType = "system";
            ContentType = "text/plain; charset=utf8";
        }
    }

    public class dfield
    {
        // Class to hold dynamic field information
        // Includes Name and Value as string properties.
        public string Name { get; set; }
        public string Value { get; set; }
    }

    public class attach
    {
        // Class to hold attachment information.
        // Includes Content, ContentType and Filename as string properties.
        public string Content { get; set; }
        public string ContentType { get; set; }
        public string Filename { get; set; }
    }

    public class TicketCreateRequest
    {
        // Class to hold the full REST request data.
        // Includes Ticket and Article as NewTicket and NewArticle properties.
        public NewTicket Ticket { get; set; }
        public NewArticle Article { get; set; }
        public List<dfield> DynamicField { get; set; }
        public List<attach> Attachment { get; set; }
   
        public TicketCreateRequest()
        {
            // Initialize NewTicket and NewArticle objects.
            Ticket = new NewTicket();
            Article = new NewArticle();
        }

        public void AddDynamicField(string _name, string _value)
        {
            // Method to add additional dynamic fields.
            if (DynamicField == null) 
            {
                // If this is the first dynamic field, initialize collection.
                DynamicField = new List<dfield>();
            }
            DynamicField.Add( new dfield() { Name = _name, Value = _value} );
        }

        public void AddAttachment(string attachfile)
        {
            // Method to add additional attachments.
            if (Attachment == null)
            {
                // If this is the first attachment, initialize collection.
                Attachment = new List<attach>();
            }
            byte[] rawcontent;
            string _content;
            string _filename;
            string _contenttype;

            rawcontent = File.ReadAllBytes(attachfile); // Read file.
            _content = Convert.ToBase64String(rawcontent);  // Convert content to base64.

            _filename = Path.GetFileName(attachfile);  // Get filename from path.
            _contenttype = MimeMapping.GetMimeMapping(_filename); // Get mimetype from fileextension.

            Attachment.Add(new attach() { Content = _content, Filename = _filename, ContentType = _contenttype });
        }
    }

    public class ResponseTicket
    {
        // Class to hold the response data from OTRS.
        // Includes TicketID, ArticleID and TicketNumer as strings.
        public string TicketID;
        public string ArticleID;
        public string TicketNumber;
    }

    public class TicketCreator
    {
        // Main TicketCreator class.
        // Use this in code to create the new ticket.
        // Includes user, password and resource as string properties
        // request as a TicketCreateRequest property
        // otrs as a HttpClient property
        // httpresponse as a HttpResponseMessage property
        // response as a ResponseTicket property
        public TicketCreateRequest request { get; set; }
        public HttpClient otrs { get; set; }
        public HttpResponseMessage httpresponse { get; set; }
        public ResponseTicket response { get; set; }

        // Define private variables for username and password, and set defaults.
        private string _user = "username";
        private string _password = "password";

        // Create write-only properties to allow access to set username and password.
        public string user
        {
            set
            {
                _user = value;
            }
        }
        public string password
        {
            set
            {
                _password = value;
            }
        }
        public string resource { get; set; }

        private UriBuilder resourceuri;
        private NameValueCollection querystring;

        public TicketCreator()
        {
            // Initialize the request, otrs, httpresponse and response objects.
            request = new TicketCreateRequest();
            otrs = new HttpClient();
            httpresponse = new HttpResponseMessage();
            response = new ResponseTicket();

            // Set default resource.
            resource = "/NewTicket";

            // Set default options for otrs object.
            otrs.BaseAddress = new Uri("http://www.wsetonlineclassroom.com/otrs/nph-genericinterface.pl/Webservice/OCAdminREST");
            otrs.DefaultRequestHeaders.Accept.Clear();
            otrs.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        // Request() method runs the REST request returning the HttpStatusCode 
        public async Task<HttpStatusCode> Request()
        {
            // Build uri from base request and resource and add username and password to query string.
            resourceuri = new UriBuilder(otrs.BaseAddress+resource);
            querystring = HttpUtility.ParseQueryString(resourceuri.Query);

            querystring["UserLogin"] = _user;
            querystring["Password"] = _password; 

            resourceuri.Query = querystring.ToString();

            // Send REST request to OTRS.
            httpresponse = await otrs.PostAsJsonAsync(resourceuri.Uri, request);

            // Process response.
            response = await httpresponse.Content.ReadAsAsync<ResponseTicket>();

            // Return the HTTP Status Code.
            return httpresponse.StatusCode;
        }

    }

}