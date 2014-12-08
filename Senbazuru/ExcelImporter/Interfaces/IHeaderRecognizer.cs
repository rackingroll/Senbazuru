using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.In4.ExcelImporter.Interfaces
{
    interface IHeaderRecognizer
    {
        /// <summary>
        /// Repair 
        /// </summary>
        /// <param name="colId">Denotes which colume to change.</param>
        /// <param name="targetRowNum">Denotes the row number of the target attribute that repaired.</param>
        /// <param name="oldParentRowNum">Denotes the row number of the target attribute’s parent attribute before repaired.</param>
        /// <param name="newParentRowNum">Denotes the row number of the target attribute’s parent attribute after repaired.</param>
        void Repair(int colId, int targetRowNum, int oldParentRowNum , int newParentRowNum);
    }
}
