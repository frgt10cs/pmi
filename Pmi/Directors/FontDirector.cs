using Pmi.Builders;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Directors
{
    class FontDirector
    {
        private ExcelFontBuilder fontBuilder;
        public ExcelFontBuilder FontBuilder { set { fontBuilder = value; } }

        public void BuildUniversityInfoFont()
        {
            fontBuilder.SetFontSize(14);
        }

        public void BuildTitleFont()
        {
            fontBuilder.SetFontSize(16);
            fontBuilder.AddBold();
        }

        public void BuildYearFont()
        {
            fontBuilder.SetFontSize(13);
            fontBuilder.AddUnderline();
        }

        public void BuildEmployeeInfoMeta()
        {
            fontBuilder.SetFontSize(8);
        }

        public void BuildEmployeeInfo()
        {
            fontBuilder.SetFontSize(14);
            fontBuilder.AddUnderline();            
        }
    }
}
