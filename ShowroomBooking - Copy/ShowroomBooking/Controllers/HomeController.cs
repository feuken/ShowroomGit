using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Exchange.WebServices.Data;
using ShowroomBooking.Models;
using ShowroomBooking.RequestObjects;

namespace ShowroomBooking.Controllers
{
    public class HomeController : Controller
    {
        private ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }


        
        //public IActionResult Booking()
        //{
        //    var mailbox = "felix.feuk@cybercom.com";

        //    service.Credentials = new WebCredentials("felix.feuk@cybercom.com", "Sommar23", "cybercom.com");
        //    service.EnableScpLookup = false;
        //    //service.AutodiscoverUrl("felix@feukit.onmicrosoft.com", RedirectionUrlValidationCallback);
        //    service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");


        //    List<AttendeeInfo> attendees = new List<AttendeeInfo>();
        //    attendees.Add(new AttendeeInfo(mailbox, MeetingAttendeeType.Required, false));
        //    //service.TraceEnabled = true;
        //    GetUserAvailabilityResults result = service.GetUserAvailability(attendees, new TimeWindow(DateTime.Now, DateTime.Now.AddDays(1)), AvailabilityData.FreeBusy);

        //    List<Events> model = new List<Events>();

            

        //    foreach (var a in result.AttendeesAvailability)
        //    {
        //        foreach (var b in a.CalendarEvents)
        //        {
        //            Events events = new Events();
        //            events.EventStart = b.StartTime;
        //            events.EventEnd = b.EndTime;

        //            model.Add(events);
        //        }
        //    }

        //    return View(model);
        //}

        //[HttpPost]
        //public IActionResult Booking(BookingRequestObjects request)
        //{
        //    service.Credentials = new WebCredentials("felix.feuk@cybercom.com", "Sommar23", "cybercom.com");
        //    service.EnableScpLookup = false;
        //    //service.AutodiscoverUrl("felix@feukit.onmicrosoft.com", RedirectionUrlValidationCallback);
        //    service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");


        //    Appointment appointment = new Appointment(service);
        //    // Set the properties on the appointment object to create the appointment.
        //    appointment.Subject = request.AppointSubject;
        //    appointment.Start = DateTime.Now; //request.AppointStart;
        //    appointment.End = DateTime.Now.AddHours(1); //request.AppointEnd;
        //    appointment.Location = "Showroom";
        //    appointment.ReminderDueBy = DateTime.Now;
        //    // Save the appointment to your calendar.
        //    appointment.Save(SendInvitationsMode.SendToNone);


        //    return RedirectToAction("Booking");
        //}

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
