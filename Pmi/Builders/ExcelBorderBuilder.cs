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

        public void AddTopBorder(string hexColorValue = "000000", BorderStyleValues borderStyle = BorderStyleValues.Thin)
        {
            border.TopBorder = new TopBorder();
            border.TopBorder.Color = new Color() { Rgb = new HexBinaryValue() { Value = hexColorValue } };
            border.TopBorder.Style = borderStyle;
        }

        public void AddBottomBorder(string hexColorValue = "000000", BorderStyleValues borderStyle = BorderStyleValues.Thin)
        {
            border.BottomBorder = new BottomBorder();
            border.BottomBorder.Color = new Color() { Rgb = new HexBinaryValue() { Value = hexColorValue } };
            border.BottomBorder.Style = borderStyle;
        }

        public void AddLeftBorder(string hexColorValue = "000000", BorderStyleValues borderStyle = BorderStyleValues.Thin)
        {
            border.LeftBorder = new LeftBorder();
            border.LeftBorder.Color = new Color() { Rgb = new HexBinaryValue() { Value = hexColorValue } };
            border.LeftBorder.Style = borderStyle;
        }

        public void AddRightBorder(string hexColorValue = "000000", BorderStyleValues borderStyle = BorderStyleValues.Thin)
        {
            border.RightBorder = new RightBorder();
            border.RightBorder.Color = new Color() { Rgb = new HexBinaryValue() { Value = hexColorValue } };
            border.RightBorder.Style = borderStyle;
        }

        public Border GetBorder()
        {
            Border resultBorder = border;
            Reset();
            return resultBorder;
        }
    }
}