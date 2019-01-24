﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Punkstar.DocHelper.Xls2Object.SampleTestAssembly
{
    public class SubEntityA
    {
        public int SubEntityAId { get; set; }
        public SubEntityB SubEntityBSample { get; set; }
        public List<SubEntityB> ListOfSubEntityBSample { get; set; }
        public string SubEntityAStringSample { get; set; }
        public List<int> ListOfIntegersSample { get; set; }
        public SubEntityA()
        {
            ListOfIntegersSample = new List<int>();
            ListOfSubEntityBSample = new List<SubEntityB>() ;
        }
        
    }
}
