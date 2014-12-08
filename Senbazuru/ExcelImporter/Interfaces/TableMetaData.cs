using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.In4.ExcelImporter.Core
{
    public class TableMetaData
    {
        public ColumnMetaData[] MetaDataOfColumns;

        public HiearchicalMetaData HiearchicalMetaData;
    }

    public class HiearchicalMetaData
    {
        List<HiearchicalTree<int>> HiearchicalTrees { get; set; }
    }

    public class HiearchicalTree<V> : HashSet<HiearchicalTree<V>>
    {
        public V Value { get; set; }
    }

    public class ColumnMetaData
    {
        public TypeCode type { get; set; }

        public bool IsDimension { get; set; }

        public bool IsMeasure { get; set; }

        public DateTimeMeta DataTimeMeta { get; set; }
    }

    public enum DateTimeMeta
    {
        Unknown,
        DateOnly,
        TimeOnly,
        DateAndTime
    }

    public enum DateMeta
    {
        Unknown,
        YearOnly,
        YearAndMonth,
        FullDate
    }

    public enum TimeMeta
    {
        Unknown,
        HourOnly,
        HourAndMinute,
        HourMinuteAndSecond,
        FullTime
    } 

    public class DateTimeMetaData
    {
        public DateTimeMeta DateTimeMeta;
        public DateMeta DateMeta;
        public TimeMeta TimeMeta;
    }
}
