using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Punkstar.DocHelper.Xls2Object.SampleTestAssembly
{
    public class MainEntity
    {
        public int MainEntityId { get; set; }
        public string StringFieldSample { get; set; }
        public DateTime DateTimeFieldSample { get; set; }
        public SubEntityA SubEntityASample { get; set; }
        public MainEntity()
        {
            SubEntityASample = new SubEntityA();
        }
    }
}
