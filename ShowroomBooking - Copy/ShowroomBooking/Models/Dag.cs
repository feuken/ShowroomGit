using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ShowroomBooking.Models
{
    public class Dag
    {   
        
        public DateTime Datum { get; set; }
        public virtual List<Tid> Tider { get; set; }
    }
}
