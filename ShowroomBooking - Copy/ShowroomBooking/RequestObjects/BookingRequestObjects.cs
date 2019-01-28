using ShowroomBooking.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;


namespace ShowroomBooking.RequestObjects
{
    public class BookingRequestObjects
    {
        [Required]
        public string Email { get; set; }
        public DateTime AppointDate { get; set; }
        public DateTime AppointStart { get; set; }
        public DateTime AppointEnd { get; set; }

    }
}
