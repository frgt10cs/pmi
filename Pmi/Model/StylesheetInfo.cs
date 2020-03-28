using Pmi.Directors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Model
{
    public class StylesheetInfo
    {        
        public Dictionary<int, uint> cellFormatIndexes { get; private set; }        

        public StylesheetInfo()
        {
            cellFormatIndexes = new Dictionary<int, uint>();
        }

        public uint GetCellFormatIndex(int cellFormatType)
        {
            return cellFormatIndexes[cellFormatType];
        }

        public void AddCellFormatIndex(int cellFormatType, uint cellFormatIndex)
        {
            cellFormatIndexes[cellFormatType] = cellFormatIndex;
        }

        public void RemoveCellFormatIndex(int cellFromatType)
        {
            cellFormatIndexes.Remove(cellFromatType);
        }
    }
}
