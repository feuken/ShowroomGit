using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ShowroomBooking.Models
{
    public class Vecka
    {
        public int VeckoNummer { get; set; }
        public virtual List<Dag> Dagar { get; set; }
    }

    
}
