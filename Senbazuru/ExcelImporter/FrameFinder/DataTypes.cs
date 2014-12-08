using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameFinder
{
    public enum DataType
    {
        General,
        Number,
        Date,
        Text,
        ZipCode,
        SocialSecurityNumber,
        PhoneNumber,
        NONE
    }
    public enum RowLabel
    {
        Title,
        Header,
        Data,
        Blank,
        Footnote
    }
    class DataTypes
    {
        public static RowLabel String2RowLabel(string s)
        {
            switch (s)
            {
                case "Title": return RowLabel.Title;
                case "Header": return RowLabel.Header;
                case "Blank": return RowLabel.Blank;
                case "Footnote": return RowLabel.Footnote;
                default: return RowLabel.Data;
            }
        }
    }
}
