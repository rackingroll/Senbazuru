using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Senbazuru.HirarchicalExtraction
{
    public class ModelFeatures
    {

        // Construction Method of ModelFeatures
        public ModelFeatures() {}

        /* Bellow are Unary Features */

        public int FeatureFirst(IList<Range> celllist, int index)
        {
            return index == 0 ? 1 : 0;
        }

        public int FeatureLast(IList<Range> celllist, int index)
        {
            return index == celllist.Count - 1 ? 1 : 0;
        }

        public int FeatureFontUnderLine(IList<Range> celllist, int index)
        {
            return celllist[index].Font.Underline == 1 ? 1 : 0;
        }

        public int FeatureFontItalic(IList<Range> celllist, int index)
        {
            return celllist[index].Font.Italic == 1 ? 1 : 0;
        }

        public int FeatureFontBold(IList<Range> celllist, int index)
        {
            return celllist[index].Font.Bold == 1 ? 1 : 0;
        }

        public int FeatureContainTotal(IList<Range> celllist, int index)
        {
            string value = (string)celllist[index].get_Value();
            return value.ToLower().Contains("total")? 1:0 ;
        }

        public int FeatureContainColon(IList<Range> celllist, int index)
        {
            string value = (string)celllist[index].get_Value();
            return value.ToLower().Contains(":") ? 1 : 0;
        }

        public int FeatureCenterAligned(IList<Range> celllist, int index)
        {
            return celllist[index].HorizontalAlignment == XlHAlign.xlHAlignCenter? 1: 0;
        }

        public int FeatureNumeric(IList<Range> celllist, int index)
        {
            string value = celllist[index].get_Value();
            string type = this.getValueType(value);

            switch (type)
            {
                case "int":
                    return 1;
                case "double":
                    return 1;
                default:
                    return 0;
            }
        }

        public int FeatureCapitalized(IList<Range> celllist, int index)
        {
            string value = (string)celllist[index].get_Value();
            for (int i = 0; i < value.Count(); i++) {
				char c = value[i];
                if (!char.IsLower(c))
                {
                    return 0;
                }
            }
            return 1;
        }

        /* Below are Binary Features */
        public int BFeatureAdjacent(IList<Range> celllist, int indexParent, int indexChild)
        {
            return Math.Abs(indexChild - indexParent) == 1 ? 1 : 0;
        }

        // Child’s indentation is greater than parent’s
        public int BFeatureChildindentationGreater(IList<Range> celllist, int indexParent, int indexChild)
        {
            string parent = celllist[indexParent].Text;
            string child = celllist[indexChild].Text;
            if (this.IndentManual(child) > this.IndentManual(parent)) return 1;
            else
                return celllist[indexParent].IndentLevel < celllist[indexChild].IndentLevel ? 1 : 0;
        }

        // Child’s row index is greater than parent’s
        public int BFeatureChildindexGreater(IList<Range> celllist, int indexParent, int indexChild)
        {
            return indexParent < indexChild ? 1 : 0;
        }

        // Child’s font size is smaller than parent’s
        public int BFeatureChildSizeSmaller(IList<Range> celllist, int indexParent, int indexChild)
        {
            if (celllist[indexParent].Font.Size is double && celllist[indexChild].Font.Size is double)
            {
                return celllist[indexParent].Font.Size > celllist[indexChild].Font.Size ? 1 : 0;
            }
            else
                return 0;
//            if (celllist[indexParent].Font.Size == System.DBNull.Value || celllist[indexChild].Font.Size == System.DBNull.Value) return 0;
            
        }

        //Has blank cells in the middle
        public int BFeatureBlankCellMiddle(IList<Range> celllist, int indexParent, int indexChild)
        {
            int indexStart = indexParent < indexChild ? indexParent : indexChild;
            int indexEnd = indexParent > indexChild ? indexParent : indexChild;

            for (int i = indexStart + 1; i < indexEnd; i++)
            {
                string cellValue = (celllist[i].Text as string);
                if (cellValue.Length == 0)
                {
                    return 1;
                }
            }
            return 0 ;
        }

        // Has middle cell with indentation between the pair’s
        public int BFeatureIndentationMiddle(IList<Range> celllist, int indexParent, int indexChild)
        {
            int indexStart = indexParent < indexChild ? indexParent : indexChild;
            int indexEnd = indexParent > indexChild ? indexParent : indexChild;
            for (int i = indexStart + 1; i < indexEnd; i++)
            {
                string parent = celllist[indexParent].Text;
                string child = celllist[indexChild].Text;
                string middle = celllist[i].Text;
                if (this.IndentManual(middle) < this.IndentManual(parent) && this.IndentManual(middle) > this.IndentManual(child)
                    || this.IndentManual(middle) > this.IndentManual(parent) && this.IndentManual(middle) < this.IndentManual(child)) return 1;
                else if (celllist[i].IndentLevel < celllist[indexParent].IndentLevel && celllist[i].IndentLevel > celllist[indexChild].IndentLevel
                    || celllist[i].IndentLevel > celllist[indexParent].IndentLevel && celllist[i].IndentLevel < celllist[indexChild].IndentLevel)
                {
                    return 1;
                }
            }
            return 0;
        }

        // Has middle cell with indentation larger the pair’s
        public int BFeatureIndentationLarger(IList<Range> celllist, int indexParent, int indexChild)
        {
            int indexStart = indexParent < indexChild ? indexParent : indexChild;
            int indexEnd = indexParent > indexChild ? indexParent : indexChild;
            for (int i = indexStart + 1; i < indexEnd; i++)
            {
                string parent = celllist[indexParent].Text;
                string child = celllist[indexChild].Text;
                string middle = celllist[i].Text;
                if (this.IndentManual(middle) > this.IndentManual(parent) && this.IndentManual(middle) > this.IndentManual(child)) return 1;
                else if (celllist[i].IndentLevel > celllist[indexParent].IndentLevel && celllist[i].IndentLevel > celllist[indexChild].IndentLevel)
                {
                    return 1;
                }
            }
            return 0;
        }

        // Has middle cell with indentation between the pair’s
        
        public int BFeatureIndentationShorter(IList<Range> celllist, int indexParent, int indexChild)
        {
            int indexStart = indexParent < indexChild ? indexParent : indexChild;
            int indexEnd = indexParent > indexChild ? indexParent : indexChild;
            for (int i = indexStart + 1; i < indexEnd; i++)
            {
                string parent = celllist[indexParent].Text;
                string child = celllist[indexChild].Text;
                string middle = celllist[i].Text;
                if (this.IndentManual(middle) < this.IndentManual(parent) && this.IndentManual(middle) < this.IndentManual(child)) return 1;
                else if (celllist[i].IndentLevel < celllist[indexParent].IndentLevel && celllist[i].IndentLevel < celllist[indexChild].IndentLevel)
                {
                    return 1;
                }
            }
            return 0;
        }

        // Has middle cell containing “:” or “total”
        public int BFeatureContainColonAndTotal(IList<Range> celllist, int indexParent, int indexChild)
        {
            int indexStart = indexParent < indexChild ? indexParent : indexChild;
            int indexEnd = indexParent > indexChild ? indexParent : indexChild;
            for (int i = indexStart + 1; i < indexEnd; i++)
            {
                string value = (string)celllist[i].Text;
                if (value == null) continue;
                if (value.ToLower().Contains("total") || value.ToLower().Contains(":"))
                {
                    return 1;
                }
            }
            return 0;
        }

        // Has Bold is different
        public int BFeatureBoldDiffer(IList<Range> celllist, int indexParent, int indexChild)
        {
            if (celllist[indexParent].Font.Bold is bool && celllist[indexChild].Font.Bold is bool)
                return celllist[indexParent].Font.Bold == celllist[indexChild].Font.Bold ? 0 : 1;
            else
                return 0;

        }

        // Has Italic is different
        public int BFeatureItalicDiffer(IList<Range> celllist, int indexParent, int indexChild)
        {
            return celllist[indexParent].Font.Italic == celllist[indexChild].Font.Italic ? 0 : 1;
        }

        // Has Underline is different
        public int BFeatureUnderlineDiffer(IList<Range> celllist, int indexParent, int indexChild)
        {
            return celllist[indexParent].Font.Underline == celllist[indexChild].Font.Underline ? 0 : 1;
        }


        /* These two features are different */
        public int BFeatureStyleAdjacent(IList<Range> celllist, int indexParent, int indexChild)
        {
            return 0;
        }

        public int BFeatureParentRoot(IList<Range> celllist, int indexParent, int indexChild)
        {
            return 0;
        }

        /* There are three features for Edge Potential*/

        // StylelisticAffinity
        public int EFeatureStylisticAffinity(IList<Range> celllist, int index1Parent,
            int index1Child, int index2Parent, int index2Child)
        {
            return celllist[index1Child].Style == celllist[index2Child].Style && celllist[index2Parent].Style == celllist[index1Parent].Style ? 1 : 0;
        }

        // MetadataAffinity, need to add domain knowledge into this feature
        // Need do
        public int EFeatureMetaDataAffinity(IList<Range> celllist, int index1Parent,
            int index1Child, int index2Parent, int index2Child)
        {
            return 0;
        }

        // AdjacentDependency need to consider what is the adjacent of this feature
        // Need do
        public int EFeatureAdjacentDependency(IList<Range> celllist, int index1Parent,
            int index1Child, int index2Parent, int index2Child)
        {
            return 0;
        }

        /*Below are auxiliary method*/
        private string getValueType(string cellValue)
        {
            int iRet;
            if (int.TryParse(cellValue, out iRet))
            {
                return "int";
            }
            double dRet;
            if (double.TryParse(cellValue, out dRet))
            {
                return "double";
            }
            return "str";
        }

        private int IndentManual(string value)
        {
            int indent = 0 ;
            foreach (char c in value)
            {
                if (c.Equals(' '))
                {
                    indent++;
                }
                else
                {
                    break;
                }
            }
            return indent;
        }
    }
}
