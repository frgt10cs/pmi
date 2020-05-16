using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using Pmi.Builders;

namespace Pmi.Directors
{
    class BorderDirector
    {
        private ExcelBorderBuilder builder;
        public ExcelBorderBuilder BorderBuilder { set { builder = value; } }

        public void BuildDefaultBorders()
        {
            builder.AddTopBorder();
            builder.AddBottomBorder();
            builder.AddRightBorder();
            builder.AddLeftBorder();
        }
    }
}