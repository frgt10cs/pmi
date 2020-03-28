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
        private IFontBuilder fontBuilder;
        public IFontBuilder FontBuilder { set { fontBuilder = value; } }

        public void BuildUniversityInfoFont()
        {
            fontBuilder.AddFontSize(14);
        }

        public void BuildTitleFont()
        {
            fontBuilder.AddFontSize(16);
            fontBuilder.AddBold();
        }

        public void BuildYearFont()
        {
            fontBuilder.AddFontSize(13);
            fontBuilder.AddUnderline();
        }

        public void BuildEmployeeInfoMeta()
        {
            fontBuilder.AddFontSize(8);
        }

        public void BuildEmployeeInfo()
        {
            fontBuilder.AddFontSize(14);
            fontBuilder.AddUnderline();            
        }
    }
}
