using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BordxGenerator.Model
{
    class Period
    {
        public int id { get; set; }
        public DateTime From { get; set; }
        public DateTime To { get; set; }
        public string Contract { get; set; }
        public string Reference { get; set; }
    }
}
