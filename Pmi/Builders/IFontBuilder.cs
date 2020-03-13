using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Builders
{
    public interface IFontBuilder
    {
        void Reset();
        void AddFontSize(int size);
        void AddFontName(string fontName);
        void AddBold();
        void AddUnderline();
        void AddColor(string hexValue);
        void AddItalic();
    }

}
