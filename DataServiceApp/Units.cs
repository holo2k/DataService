using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Linq.Mapping;

namespace DataServiceApp
{
    [Table(Name ="DBMAIN")]
    class Units
    {
        [Column(Name = "DateTime")]
        public DateTime DateTime { get; set; }

        [Column(Name = "Index")]
        public int Index { get; set; }
    }
}
