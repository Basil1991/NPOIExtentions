﻿using Permaisuri.Tools.NPOIExtentions.Tests;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Permaisuri.Tools.NPOIExtentionsTests {
    class Program {
        static void Main() {
            ExcelHelpTests abc = new ExcelHelpTests();
            abc.GetXlsxByDt();
            //abc.GetXlsxByDs();
            //abc.GetXlsxDynamic();
            //abc.GetXlsxByDynamicList();
            //abc.GetXlsByDt();
            //abc.GetXlsByDs();
            //abc.GetXlsDynamic();
            //abc.GetXlsByDynamicList();
        }
    }
}
