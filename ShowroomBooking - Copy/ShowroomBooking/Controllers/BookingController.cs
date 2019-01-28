using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Exchange.WebServices.Data;
using Newtonsoft.Json;
using ShowroomBooking.Models;
using ShowroomBooking.RequestObjects;

namespace ShowroomBooking.Controllers
{
 
    public class BookingController : Controller
    {
        private ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);

        public IActionResult Index()
        {
            var mailbox = "felix.feuk@cybercom.com";

            service.Credentials = new WebCredentials("felix.feuk@cybercom.com", "Sommar23", "cybercom.com");
            service.EnableScpLookup = false;
            //service.AutodiscoverUrl("felix@feukit.onmicrosoft.com", RedirectionUrlValidationCallback);
            service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");


            List<AttendeeInfo> attendees = new List<AttendeeInfo>();
            attendees.Add(new AttendeeInfo(mailbox, MeetingAttendeeType.Required, false));
            //service.TraceEnabled = true;
            GetUserAvailabilityResults result = service.GetUserAvailability(attendees, new TimeWindow(DateTime.Now, DateTime.Now.AddDays(1)), AvailabilityData.FreeBusy);

            List<Events> model = new List<Events>();

            foreach (var a in result.AttendeesAvailability)
            {
                foreach (var b in a.CalendarEvents)
                {
                    Events events = new Events();
                    events.EventStart = b.StartTime;
                    events.EventEnd = b.EndTime;
                    

                    model.Add(events);
                }
            }

            return View(model);
        }

        public List<Tid> GetTider()
        {
            List<Tid> Tider = new List<Tid>();

            Tider.Add(new Tid() { StartTime = DateTime.Parse("12/12/2012 09:00:00"), EndTime = DateTime.Parse("12/12/2012 10:00:00") });
            Tider.Add(new Tid() { StartTime = DateTime.Parse("12/12/2012 10:00:00"), EndTime = DateTime.Parse("12/12/2012 11:00:00") });
            Tider.Add(new Tid() { StartTime = DateTime.Parse("12/12/2012 11:00:00"), EndTime = DateTime.Parse("12/12/2012 12:00:00") });
            Tider.Add(new Tid() { StartTime = DateTime.Parse("12/12/2012 12:00:00"), EndTime = DateTime.Parse("12/12/2012 13:00:00") });
            Tider.Add(new Tid() { StartTime = DateTime.Parse("12/12/2012 13:00:00"), EndTime = DateTime.Parse("12/12/2012 14:00:00") });
            Tider.Add(new Tid() { StartTime = DateTime.Parse("12/12/2012 14:00:00"), EndTime = DateTime.Parse("12/12/2012 15:00:00") });
            Tider.Add(new Tid() { StartTime = DateTime.Parse("12/12/2012 15:00:00"), EndTime = DateTime.Parse("12/12/2012 16:00:00") });
            Tider.Add(new Tid() { StartTime = DateTime.Parse("12/12/2012 16:00:00"), EndTime = DateTime.Parse("12/12/2012 17:00:00") });
            Tider.Add(new Tid() { StartTime = DateTime.Parse("12/12/2012 17:00:00"), EndTime = DateTime.Parse("12/12/2012 18:00:00") });

            return Tider;
        }

        public Vecka GetData()
        {
            var idag = DateTime.Now;
            
            Vecka vecka1 = new Vecka();
            vecka1.VeckoNummer = 1;
            Dag måndag = new Dag();
            Dag tisdag = new Dag();
            Dag onsdag = new Dag();
            Dag torsdag = new Dag();
            Dag fredag = new Dag();


            måndag.Tider = GetTider();
            tisdag.Tider = GetTider();
            onsdag.Tider = GetTider();
            torsdag.Tider = GetTider();
            fredag.Tider = GetTider();

            int dayOfWeek = (int)DateTime.Now.DayOfWeek;
            if (dayOfWeek == 1)
            {
                måndag.Datum = idag.Date;
                tisdag.Datum = idag.Date.AddDays(1);
                onsdag.Datum = idag.Date.AddDays(2);
                torsdag.Datum = idag.Date.AddDays(3);
                fredag.Datum = idag.Date.AddDays(4);

            }
            else if (dayOfWeek == 2)
            {
                måndag.Datum = idag.Date.AddDays(-1);
                tisdag.Datum = idag.Date.AddDays(0);
                onsdag.Datum = idag.Date.AddDays(1);
                torsdag.Datum = idag.Date.AddDays(2);
                fredag.Datum = idag.Date.AddDays(3);

            }
            else if (dayOfWeek == 3)
            {
                måndag.Datum = idag.Date.AddDays(-2);
                tisdag.Datum = idag.Date.AddDays(-1);
                onsdag.Datum = idag.Date.AddDays(0);
                torsdag.Datum = idag.Date.AddDays(1);
                fredag.Datum = idag.Date.AddDays(2);

            }


            if (dayOfWeek == 4)
            {
                måndag.Datum = idag.Date.AddDays(-3);
                tisdag.Datum = idag.Date.AddDays(-2);
                onsdag.Datum = idag.Date.AddDays(-1);
                torsdag.Datum = idag.Date;
                fredag.Datum = idag.Date.AddDays(1);

            }

            if (dayOfWeek == 5)
            {
                måndag.Datum = idag.Date.AddDays(-4);
                tisdag.Datum = idag.Date.AddDays(-3);
                onsdag.Datum = idag.Date.AddDays(-2);
                torsdag.Datum = idag.Date.AddDays(-1);
                fredag.Datum = idag.Date;

            }

            vecka1.Dagar = new List<Dag> { måndag, tisdag, onsdag, torsdag, fredag };

            return vecka1;
        }

        public IActionResult Booking()
        {
            Vecka vecka1 = GetData();

            List<Events> events = GetEvents();

            foreach (Dag dag in vecka1.Dagar)
            {

                foreach (Tid tid in dag.Tider)
                {
                    foreach (Events evnt in events)
                    {
                        if (evnt.EventStart == new DateTime(dag.Datum.Year, dag.Datum.Month, dag.Datum.Day, tid.StartTime.Hour, tid.StartTime.Minute,0) && evnt.EventEnd == new DateTime(dag.Datum.Year, dag.Datum.Month, dag.Datum.Day, tid.EndTime.Hour, tid.EndTime.Minute, 0))
                        {
                            tid.Booked = true;
                        }
                    }
                }
                
            }

            return View(vecka1);
        }



        public ActionResult Appointment(string i, string j)
        {

            Vecka vecka = GetData();
            BookingRequestObjects obj = new BookingRequestObjects();
            obj.AppointStart = vecka.Dagar[(Convert.ToInt32(i))].Tider[(Convert.ToInt32(j))].StartTime;
            obj.AppointEnd = vecka.Dagar[(Convert.ToInt32(i))].Tider[(Convert.ToInt32(j))].EndTime;
            obj.AppointDate = vecka.Dagar[(Convert.ToInt32(i))].Datum;


            return View(obj);
        }

        [HttpPost]
        public IActionResult Appointment(BookingRequestObjects request)
        {

            service.Credentials = new WebCredentials("felix.feuk@cybercom.com", "Sommar23", "cybercom.com");
            service.EnableScpLookup = false;
            //service.AutodiscoverUrl("felix@feukit.onmicrosoft.com", RedirectionUrlValidationCallback);
            service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");

            Appointment meeting = new Appointment(service);
            // Set the properties on the meeting object to create the meeting.
            meeting.Subject = "Introduction to Cybercom Showroom";
            meeting.Body = "Welcome to a introduction of the Cybercom Showroom!";
            
            meeting.Start =new DateTime(request.AppointDate.Year, request.AppointDate.Month, request.AppointDate.Day ,
                                        request.AppointStart.Hour,request.AppointStart.Minute,0);
            meeting.End = new DateTime(request.AppointDate.Year, request.AppointDate.Month, request.AppointDate.Day,
                                        request.AppointEnd.Hour, request.AppointEnd.Minute, 0); ;
            meeting.Location = "Showroom";
            meeting.RequiredAttendees.Add(request.Email);
            meeting.AllowNewTimeProposal = false;

            meeting.ReminderMinutesBeforeStart = 15;
            // Save the meeting to the Calendar folder and send the meeting request.
            meeting.Save(SendInvitationsMode.SendToAllAndSaveCopy);

            TempData["custdetails"] = "Booking is confirmed";

            Thread.Sleep(100);
            return RedirectToAction("Booking");
        }

        public List<Events> GetEvents()
        {
            var mailbox = "felix.feuk@cybercom.com";

            service.Credentials = new WebCredentials("felix.feuk@cybercom.com", "Sommar23", "cybercom.com");
            service.EnableScpLookup = false;
            //service.AutodiscoverUrl("felix@feukit.onmicrosoft.com", RedirectionUrlValidationCallback);
            service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");


            List<AttendeeInfo> attendees = new List<AttendeeInfo>();
            attendees.Add(new AttendeeInfo(mailbox, MeetingAttendeeType.Required, false));
            //service.TraceEnabled = true;
            GetUserAvailabilityResults result = service.GetUserAvailability(attendees, new TimeWindow(DateTime.Now, DateTime.Now.AddDays(1)), AvailabilityData.FreeBusy);

            List<Events> model = new List<Events>();

            foreach (var a in result.AttendeesAvailability)
            {
                foreach (var b in a.CalendarEvents)
                {
                    Events events = new Events();
                    events.EventStart = b.StartTime;
                    events.EventEnd = b.EndTime;


                    model.Add(events);
                }
            }

            return model;
        }
    }
}