using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Pmi.Builders
{
    class ExcelBorderBuilder
    {
        private Border border;

        public ExcelBorderBuilder()
        {
            Reset();
        }

        public void Reset()
        {
            border = new Border();
        }

        public void GetBorder()
        {
            border.TopBorder = new TopBorder()
            {
                Color = new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                Style = BorderStyleValues.Thin
            };
            border.BottomBorder = new BottomBorder
            {
                Color = new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                Style = BorderStyleValues.Thin
            };
            border.LeftBorder = new LeftBorder
            {
                Color = new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                Style = BorderStyleValues.Thin
            };
            border.RightBorder = new RightBorder
            {
                Color = new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                Style = BorderStyleValues.Thin
            };
        }

        public Border GetBorders()
        {
            var border = this.border;
            Reset();
            return border;
        }
    }
}