using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Punkstar.DocHelper.Xls2Object.SampleTestAssembly
{
    public class MainEntity
    {
        int MainEntityId { get; set; }
        string StringFieldSample { get; set; }
        DateTime DateTimeFieldSample { get; set; }
        SubEntityA SubEntityASample { get; set; }
    }
}
