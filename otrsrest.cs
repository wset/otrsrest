using System;
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
using System.Runtime.InteropServices;
using System.Security.Cryptography;


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
            CustomerUser = Properties.otrsrest.Default.customer;
            Queue = Properties.otrsrest.Default.queue;
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

    [ComVisible(true)]
    [Guid("76FC70C9-A9C6-49BF-A444-0CE5D212BB2C")]
    public interface IUpdateSettings
    {
        bool Update(string name, string value);
        string Get(string name);
    }

    [ComVisible(true)]
    [Guid("86353DF1-8530-495D-B8A3-02281CC5A1F5")]
    [ClassInterface(ClassInterfaceType.None)]
    public class UpdateSettings : IUpdateSettings
    {
        [System.STAThread]
        public bool Update(string _name, string _value)
        {
        // Function to update stored login details for the OTRS connector.

            if (_name == "password")
            {
                // Encode passwords
                byte[] unenc = Encoding.Unicode.GetBytes(_value);

                // Generate additional entropy (will be used as the Initialization vector)
                byte[] entropy = new byte[20];
                using (RNGCryptoServiceProvider rng = new RNGCryptoServiceProvider())
                {
                    rng.GetBytes(entropy);
                }

                byte[] enc = ProtectedData.Protect(unenc, entropy, DataProtectionScope.CurrentUser);

                Properties.otrsrest.Default.password = enc;
                Properties.otrsrest.Default.entropy = entropy;
                Properties.otrsrest.Default.Save();
                return true;
            }
            else if( _name != "entropy" && Properties.otrsrest.Default[_name] != null )
            {
                // If setting exists update it.
                Properties.otrsrest.Default[_name] = _value;
                Properties.otrsrest.Default.Save();
                return true;
            }
            else 
            {
                return false;
            }
        }

        [System.STAThread]
        public string Get(string _name)
        {
            if(_name != "password" && _name != "entropy" && Properties.otrsrest.Default[_name] != null )
            {
                return Properties.otrsrest.Default[_name].ToString();
            }
            else
            {
                return "";
            }
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
        private string _user;
        private string _password;

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
            resource = "/" + Properties.otrsrest.Default.resource;

            // Set default options for otrs object.
            otrs.BaseAddress = new Uri(Properties.otrsrest.Default.uri);
            otrs.DefaultRequestHeaders.Accept.Clear();
            otrs.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            // Load username and password.
            _user = Properties.otrsrest.Default.user;
            byte[] enc = Properties.otrsrest.Default.password;
            byte[] entropy = Properties.otrsrest.Default.entropy;
            byte[] unenc = ProtectedData.Unprotect(enc, entropy, DataProtectionScope.CurrentUser);
            _password = Encoding.Unicode.GetString(unenc);
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

    [ComVisible(true)]
    [Guid("6EA65648-9201-4AF8-ACC4-CD3B57C43619")]
    public interface ICreateNewTicket
    {
        string CreateTicket(string Subject, string Message, [Optional] string Title, [Optional] string Customer, [Optional] string Queue, [Optional] string State, [Optional] int Priority, [Optional] string Attachment);
    }

    [ComVisible(true)]
    [Guid("1793697B-13FB-442A-8AB6-76781119EA4D")]
    [ClassInterface(ClassInterfaceType.None)]
    public class CreateNewTicket : ICreateNewTicket
    {
        private TicketCreator myticket;
        private string TicketNumber;
        private string TicketID;
        private string ArticleID;

        public CreateNewTicket()
        {
            
            myticket = new TicketCreator();
        }

        [System.STAThread]
        public string CreateTicket(string Subject, string Message, [Optional] string Title, [Optional] string Customer, [Optional] string Queue, [Optional] string State, [Optional] int Priority, [Optional] string Attachment)
        {
            if (!String.IsNullOrEmpty(Customer))
            {
                myticket.request.Ticket.CustomerUser = Customer;
            }
            if (!String.IsNullOrEmpty(Queue))
            {
                myticket.request.Ticket.Queue = Queue;
            }
            if (!String.IsNullOrEmpty(State))
            {
                myticket.request.Ticket.State = State;
            }
            if (Priority != 0)
            {
                myticket.request.Ticket.PriorityID = Priority;
            }
            if (!String.IsNullOrEmpty(Attachment))
            {
                myticket.request.AddAttachment(Attachment);
            }

            if (String.IsNullOrEmpty(Title)) {
                Title = Subject;
            }
            if (String.IsNullOrEmpty(Subject)) {
                Subject = Title;
            }

            myticket.request.Ticket.Title = Title;
            myticket.request.Article.Subject = Subject;
            myticket.request.Article.Body = Message;

            myticket.Request().Wait();

            TicketNumber = myticket.response.TicketNumber;
            TicketID = myticket.response.TicketID;
            ArticleID = myticket.response.ArticleID;

            return TicketNumber;
        }
    }

    [ComVisible(true)]
    [Guid("4FF8D473-99F3-479F-B90F-781DB6249D44")]
    public interface Ibuildversion
    {
        string Get();
        int major();
        int minor();
        int buildno();
        int revision();
    }

    [ComVisible(true)]
    [Guid("CF461257-D6D4-4CB1-9210-FB6143A95C1C")]
    [ClassInterface(ClassInterfaceType.None)]
    public class buildversion : Ibuildversion
    {
        public buildversion()
        {
        }

        [System.STAThread]
        public string Get()
        {
            return "v" + typeof(CreateNewTicket).Assembly.GetName().Version;
        }

        [System.STAThread]
        public int major()
        {
            return typeof(CreateNewTicket).Assembly.GetName().Version.Major;
        }

        [System.STAThread]
        public int minor()
        {
            return typeof(CreateNewTicket).Assembly.GetName().Version.Minor;
        }

        [System.STAThread]
        public int buildno()
        {
            return typeof(CreateNewTicket).Assembly.GetName().Version.Build;
        }

        [System.STAThread]
        public int revision()
        {
            return typeof(CreateNewTicket).Assembly.GetName().Version.Revision;
        }
    }

}
