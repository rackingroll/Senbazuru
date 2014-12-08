using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.In4.ExcelImporter.Core
{
    public class FlatTable
    {
        /// <summary>
        /// The title of the flat table
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Column count of the flat table
        /// </summary>
        public int ColumnCount { get; set; }

        /// <summary>
        /// Row count of the flat table exlcuding column row
        /// </summary>
        public int RowCount { get; set; }

        /// <summary>
        /// The column names of the flat table
        /// </summary>
        public string[] ColumnNames { get; set; }

        /// <summary>
        /// The column meta data of the table
        /// </summary>
        public TableMetaData MetaData { get; set; }

        /// <summary>
        /// The values of the data
        /// Consumer of the data need to convert the data into proper data type according table meta data
        /// </summary>
        public string[][] Data { get; set; }
    }
}
