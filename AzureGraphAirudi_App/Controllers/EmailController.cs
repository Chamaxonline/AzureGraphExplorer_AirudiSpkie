using Azure.Core;
using Azure.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;

namespace AzureGraphAirudi_App.Controllers
{
    public class EmailController : Controller
    {
        private readonly IConfiguration _configuration;
        public EmailController(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        public IActionResult Index()
        {
            return View();
        }


        public JsonResult Users()
        {
            try
            {
                var identity = this.User.Identity as ClaimsIdentity;
                var userId = identity.Claims.FirstOrDefault(x => x.Type == "http://schemas.microsoft.com/identity/claims/objectidentifier")?.Value;

                var clientId = _configuration.GetValue<string>("AzureAd:ClientId");
                var tenantId = _configuration.GetValue<string>("AzureAd:TenantId");
                var clientSecret = _configuration.GetValue<string>("AzureAd:ClientSecret");
                var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);
                GraphServiceClient graphServiceClient = new GraphServiceClient(clientSecretCredential);
                 var users = graphServiceClient.Users.Request().Select(x=> x.Calendar).GetAsync().Result;
                var logs = graphServiceClient.AuditLogs.SignIns.Request().GetAsync().Result;
                

               // var scopes = new[] { "https://graph.microsoft.com/.default" };
                //var tokenRequestContext = new TokenRequestContext(scopes);
                //var token = clientSecretCredential.GetTokenAsync(tokenRequestContext).Result.Token;
                // graphServiceClient.Users.Request().Header("Authorization","Bearer "+ token);

                //var calandar2 = graphServiceClient.Users["a01b941b-5745-4836-be55-ff28642f089b"].Calendars.Request().GetAsync().Result;
                var events = graphServiceClient.Users[userId].Events.Request().GetAsync().Result;
                return Json(events);
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        public JsonResult FintTime()
        {
            try
            {
                var claimsIdentity = (ClaimsIdentity)User.Identity;
                var claims = claimsIdentity.FindFirst(ClaimTypes.NameIdentifier);

                var clientId = _configuration.GetValue<string>("AzureAd:ClientId");
                var tenantId = _configuration.GetValue<string>("AzureAd:TenantId");
                var clientSecret = _configuration.GetValue<string>("AzureAd:ClientSecret");
                var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);
                GraphServiceClient graphServiceClient = new GraphServiceClient(clientSecretCredential);

                var attendees = new List<AttendeeBase>()
{
    new AttendeeBase
    {
        Type = AttendeeType.Required,
        EmailAddress = new EmailAddress
        {
            Name = "Adele Vance",
            Address = "AdeleV@8j4tv4.onmicrosoft.com"
        }
    }
};

                var locationConstraint = new LocationConstraint
                {
                    IsRequired = false,
                    SuggestLocation = false,
                    Locations = new List<LocationConstraintItem>()
    {
        new LocationConstraintItem
        {
            ResolveAvailability = false,
            DisplayName = "Conf room Hood"
        }
    }
                };

                var timeConstraint = new TimeConstraint
                {
                    ActivityDomain = ActivityDomain.Work,
                    TimeSlots = new List<TimeSlot>()
    {
        new TimeSlot
        {
            Start = new DateTimeTimeZone
            {
                DateTime = "2022-11-23T09:00:00",
                TimeZone = "Pacific Standard Time"
            },
            End = new DateTimeTimeZone
            {
                DateTime = "2022-11-25T17:00:00",
                TimeZone = "Pacific Standard Time"
            }
        }
    }
                };

                var isOrganizerOptional = false;

                var meetingDuration = new Duration("PT1H");

                var returnSuggestionReasons = true;

                var minimumAttendeePercentage = (double)100;


                var events =  graphServiceClient.Users[claims.Value]
                    .FindMeetingTimes(attendees, locationConstraint, timeConstraint, meetingDuration, null, isOrganizerOptional, returnSuggestionReasons, minimumAttendeePercentage)
                    .Request()
                    .Header("Prefer", "outlook.timezone=\"Pacific Standard Time\"")
                  
                    .PostAsync().Result;

              //  var calandar2 = graphServiceClient.Users["a01b941b-5745-4836-be55-ff28642f089b"].Calendars.Request().GetAsync().Result;
                return Json(events);
            }
            catch (Exception ex)
            {

                throw;
            }
        }
    }
}
