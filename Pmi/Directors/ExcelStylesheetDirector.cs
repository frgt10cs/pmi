using Pmi.Builders;
using Pmi.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

enum ExcelCellFormat
{
    UniveristyInfo = 0,
    Title = 1
}

namespace Pmi.Directors
{    
    public class ExcelStylesheetDirector
    {
        private ExcelStylesheetBuilder stylesheetBuilder;
        public ExcelStylesheetBuilder StylesheetBuilder { set { stylesheetBuilder = value; } }
        private FontDirector fontDirector;
        private ExcelFontBuilder fontBuilder;
        private ExcelCellFormatBuilder cellFormatBuilder;
        private ExcelCellFormatDirector cellFormatDirector;

        public ExcelStylesheetDirector()
        {
            //stylesheetBuilder = new ExcelStylesheetBuilder();
            fontBuilder = new ExcelFontBuilder();
            fontDirector = new FontDirector() { FontBuilder = fontBuilder };
            cellFormatBuilder = new ExcelCellFormatBuilder();
            cellFormatDirector = new ExcelCellFormatDirector() { CellFormatBuilder = cellFormatBuilder };
        }

        public StylesheetInfo BuildReportStylesheet()
        {
            var info = new StylesheetInfo();
            fontDirector.BuildUniversityInfoFont();
            cellFormatDirector.BuildUniveristyInfoCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            info.AddCellFormatIndex((int)ExcelCellFormat.UniveristyInfo, stylesheetBuilder.AddCellFormat(cellFormatBuilder.GetCellFormat()));

            fontDirector.BuildTitleFont();
            cellFormatDirector.BuildTitleCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            info.AddCellFormatIndex((int)ExcelCellFormat.Title, stylesheetBuilder.AddCellFormat(cellFormatBuilder.GetCellFormat()));

            return info;
        }
    }

}
