using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

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
            border.TopBorder = new TopBorder();
            border.TopBorder.Color = new Color() {  Rgb = new HexBinaryValue() { Value = "000000" }  };
            border.TopBorder.Style = BorderStyleValues.Thin;
            border.BottomBorder = new BottomBorder();
            border.BottomBorder.Color = new Color() { Rgb = new HexBinaryValue() { Value = "000000" } };
            border.BottomBorder.Style = BorderStyleValues.Thin;
            border.LeftBorder = new LeftBorder();
            border.LeftBorder.Color = new Color() { Rgb = new HexBinaryValue() { Value = "000000" } };
            border.LeftBorder.Style = BorderStyleValues.Thin;
            border.RightBorder = new RightBorder();
            border.RightBorder.Color = new Color() { Rgb = new HexBinaryValue() { Value = "000000" } };
            border.RightBorder.Style = BorderStyleValues.Thin;
        }

        public Border GetBorders()
        {
            Border border = this.border;
            Reset();
            return border;
        }
    }
}