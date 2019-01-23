using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Punkstar.DocHelper.Xls2Object
{
    public class ValidationResult
    {
        public bool Status { get; set; }
        public List<string> Messages { get; set; }
        public ValidationResult()
        {
            Status = true;
            Messages = new List<string>();
        }

    }
}
