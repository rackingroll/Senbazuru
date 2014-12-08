using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace FrameFinder
{
    class SheetLoader
    {
        private static Application APP = new Application();
        private Workbook workbook;
        public SheetLoader()
        {
            this.workbook = null;
        }
        public SheetLoader(String aFilePath)
        {
            this.workbook = APP.Workbooks.Open(Path.GetFullPath(aFilePath));
        }

        public void InitWorkbook(String aFilePath)
        {
            if (this.workbook != null)
            {
                this.workbook.Close();
            }
            this.workbook = APP.Workbooks.Open(Path.GetFullPath(aFilePath));
        }

        public void CloseWorkbook()
        {
            this.workbook.Close();
            Marshal.ReleaseComObject(this.workbook);
            this.workbook = null;
        }

        public IEnumerable<Tuple<string, MSheet>> FetchSheetDict()
        {
            Sheets sheets = this.workbook.Sheets;
            foreach (Worksheet sheet in sheets)
            {
                string sheetName = sheet.Name;
                List<Tuple<int, int, int, int>> ranges = TableHelper.SplitTable(sheet);
                foreach (Tuple<int, int, int, int> range in ranges)
                {
                    MSheet mSheet = new MSheet();
                    int upRow = range.Item1;
                    int leftCol = range.Item2;
                    int downRow = range.Item3;
                    int rightCol = range.Item4;
                    int rowNum = downRow - upRow + 1;
                    int colNum = rightCol - leftCol + 1;
                    mSheet.StartRow = upRow;
                    mSheet.StartCol = leftCol;
                    mSheet.RowNum = rowNum;
                    mSheet.ColNum = colNum;
                    {
                        Range beginCell = sheet.Cells[upRow, leftCol];
                        Range endCell = sheet.Cells[downRow, rightCol];
                        string beginAddress = beginCell.get_Address().Replace("$", "");
                        string endAddress = endCell.get_Address().Replace("$", "");
                        mSheet.Cells = sheet.get_Range(beginAddress, endAddress);
                    }
                    bool[,] vis = new bool[rowNum, colNum];
                    DataType[,] typeTable = new DataType[rowNum, colNum];
                    for (int i = 0; i < rowNum; ++i)
                    {
                        for (int j = 0; j < colNum; ++j)
                        {
                            vis[i, j] = false;
                            typeTable[i, j] = DataType.NONE;
                        }
                    }

                    for (int row = upRow; row <= downRow; ++row)
                    {
                        for (int col = leftCol; col <= rightCol; ++col)
                        {
                            if (vis[row - upRow, col - leftCol])
                            {
                                continue;
                            }
                            Range cell = sheet.Cells[row, col];
                            if (cell.MergeCells)
                            {
                                int rowCount = cell.MergeArea.Rows.Count;
                                int colCount = cell.MergeArea.Columns.Count;
                                mSheet.AddMergeCell(row, row + rowCount - 1, col, col + colCount - 1);
                                for (int i = row; i < row + rowCount; ++i)
                                {
                                    for (int j = col; j < col + colCount; ++j)
                                    {
                                        vis[i - upRow, j - leftCol] = true;
                                    }
                                }
                            }
                            string cellValue = Convert.ToString(cell.Value2);
                            string cellType = (cell.NumberFormat as string);
                            if (cellValue == null || cellValue.Length == 0)
                            {
                                continue;
                            }
                            string cType = this.getValueType(cellValue);
                            int indents = cell.IndentLevel;
                            /* XlHAlign
                             * -4131 = xlHAlignLeft                  -> ALIGN_LEFT = 0x1
                             * -4152 = xlHAlignRight                 -> ALIGN_RIGHT = 0x3
                             * -4108 = xlHAlignCenter                -> ALIGN_CENTER = 0x2
                             * -4130 = xlHAlignJustify               -> ALIGN_JUSTIFY = 0x5
                             * -4117 = xlHAlignDistributed           ->
                             * 1     = xlHAlignGeneral               -> ALIGN_GENERAL = 0x0
                             * 5     = xlHAlignFill                  -> ALIGN_FILL = 0x4
                             * 7     = xlHAlignCenterAcrossSelection ->
                             */
                            // int alignStyle = cell.HorizontalAlignment;
                            // int alignStyle = this.getFeatureAlignStyle(cell.HorizontalAlignment);
                            int alignStyle = cell.HorizontalAlignment;
                            /* XlLineStyle
                             * -4142 = xlLineStyleNone -> BORDER_NONE = 0x0
                             * -4119 = xlDouble        -> BORDER_DOUBLE = 0x6
                             * -4118 = xlDot           -> BORDER_HAIR = 0x7
                             * -4115 = xlDash          -> BORDER_DASHED = 0x3
                             * 1     = xlContinuous
                             * 4     = xlDashDot       -> BORDER_DASH_DOT = 0x9
                             * 5     = xlDashDotDot    -> BORDER_DASH_DOT_DOT = 0xB
                             * 13    = xlSlantDashDot  -> BORDER_SLANTED_DASH_DOT = 0xD
                             */
                            string borderStyle = this.getFeatureBorderStyle(cell.Borders);
                            /* XlColorIndex
                             * -4142 = xlColorIndexNone
                             * -4105 = xlColorIndexAutomatic
                             */
                            double bgColor = cell.Interior.ColorIndex;
                            int boldFlag = this.getFeatureFontBold(cell.Font);
                            double height = this.getFeatureFontHeight(cell.Font) * 20.0;
                            int italicFlag = this.getFeatureFontItalic(cell.Font);
                            // XlUnderlineStyle
                            int underlineFlag = this.getFeatureFontUnderline(cell.Font);
                            DataType dataType = this.getDataType(cell.NumberFormat as string, cellValue);
                            mSheet.InsertCell(row, col, cType, indents, alignStyle, borderStyle,
                                (int)bgColor, boldFlag, (int)height, italicFlag, underlineFlag, cellValue);
                            typeTable[row - upRow, col - leftCol] = dataType;
                        }
                    }
                    DataType[] columnTypes = new DataType[colNum];
                    this.findColumnType(typeTable, rowNum, colNum, columnTypes);
                    mSheet.SetColumnTypeTable(columnTypes, colNum);
                    yield return new Tuple<string, MSheet>(sheetName, mSheet);
                }
            }
        }

        // This method is out of date and should not be used!
        private Dictionary<string, MSheet> LoadSheetDictByTransposition()
        {
            Dictionary<string, MSheet> sheetDict = new Dictionary<string, MSheet>();
            Sheets sheets = this.workbook.Sheets;
            foreach (Worksheet sheet in sheets)
            {
                // List<Tuple<int, int, int, int>> ranges = TableHelper.SplitTable(sheet);
                string sheetName = sheet.Name;
                MSheet mSheet = new MSheet();
                int rowNum = sheet.UsedRange.Cells.Rows.Count;
                int colNum = sheet.UsedRange.Cells.Columns.Count;
                int stRow = sheet.UsedRange.Row;
                int stCol = sheet.UsedRange.Column;
                mSheet.StartRow = stRow;
                mSheet.StartCol = stCol;
                mSheet.RowNum = rowNum;
                mSheet.ColNum = colNum;
                bool[,] vis = new bool[rowNum, colNum];
                for (int i = 0; i < rowNum; ++i)
                {
                    for (int j = 0; j < colNum; ++j)
                    {
                        vis[i, j] = false;
                    }
                }

                for (int row = stRow; row < stRow + rowNum; ++row)
                {
                    for (int col = stCol; col < stCol + colNum; ++col)
                    {
                        Range cell = sheet.Cells[row, col];

                        if (vis[row - stRow, col - stCol])
                        {
                            continue;
                        }

                        if (cell.MergeCells)
                        {
                            int rowCount = cell.MergeArea.Rows.Count;
                            int colCount = cell.MergeArea.Columns.Count;
                            mSheet.AddMergeCell(col, col + colCount - 1, row, row + rowCount - 1);
                            for (int i = row; i < row + rowCount; ++i)
                            {
                                for (int j = col; j < col + colCount; ++j)
                                {
                                    vis[i - stRow, j - stCol] = true;
                                }
                            }
                        }
                        string cellValue = Convert.ToString(cell.Value2);
                        string cellType = (cell.NumberFormat as string);
                        if (cellValue == null || cellValue.Length == 0)
                        {
                            continue;
                        }
                        /*
                         * 0.00 means precision is 2
                         * #,## means to use , delimiter
                         */
                        string cType = this.getValueType(cellValue);
                        string cStr = cellValue;
                        int indents = cell.IndentLevel;
                        /* XlHAlign
                         * -4131 = xlHAlignLeft                  -> ALIGN_LEFT = 0x1
                         * -4152 = xlHAlignRight                 -> ALIGN_RIGHT = 0x3
                         * -4108 = xlHAlignCenter                -> ALIGN_CENTER = 0x2
                         * -4130 = xlHAlignJustify               -> ALIGN_JUSTIFY = 0x5
                         * -4117 = xlHAlignDistributed           ->
                         * 1     = xlHAlignGeneral               -> ALIGN_GENERAL = 0x0
                         * 5     = xlHAlignFill                  -> ALIGN_FILL = 0x4
                         * 7     = xlHAlignCenterAcrossSelection ->
                         */
                        // int alignStyle = cell.HorizontalAlignment;
                        // int alignStyle = this.getFeatureAlignStyle(cell.HorizontalAlignment);
                        int alignStyle = cell.HorizontalAlignment;
                        /* XlLineStyle
                         * -4142 = xlLineStyleNone -> BORDER_NONE = 0x0
                         * -4119 = xlDouble        -> BORDER_DOUBLE = 0x6
                         * -4118 = xlDot           -> BORDER_HAIR = 0x7
                         * -4115 = xlDash          -> BORDER_DASHED = 0x3
                         * 1     = xlContinuous
                         * 4     = xlDashDot       -> BORDER_DASH_DOT = 0x9
                         * 5     = xlDashDotDot    -> BORDER_DASH_DOT_DOT = 0xB
                         * 13    = xlSlantDashDot  -> BORDER_SLANTED_DASH_DOT = 0xD
                         */
                        string borderStyle = this.getFeatureBorderStyle(cell.Borders);
                        /* XlColorIndex
                         * -4142 = xlColorIndexNone
                         * -4105 = xlColorIndexAutomatic
                         */
                        double bgColor = cell.Interior.ColorIndex;
                        int boldFlag = this.getFeatureFontBold(cell.Font);
                        double height = this.getFeatureFontHeight(cell.Font) * 20.0;
                        int italicFlag = this.getFeatureFontItalic(cell.Font);
                        // XlUnderlineStyle
                        int underlineFlag = this.getFeatureFontUnderline(cell.Font);
                        mSheet.InsertCell(col, row, cType, indents, alignStyle, borderStyle,
                            (int)bgColor, boldFlag, (int)height, italicFlag, underlineFlag, cStr);
                    }
                }
                sheetDict.Add(sheetName, mSheet);
            }
            return sheetDict;
        }

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

        /* XlHAlign
         * -4131 = xlHAlignLeft                  -> ALIGN_LEFT = 0x1
         * -4152 = xlHAlignRight                 -> ALIGN_RIGHT = 0x3
         * -4108 = xlHAlignCenter                -> ALIGN_CENTER = 0x2
         * -4130 = xlHAlignJustify               -> ALIGN_JUSTIFY = 0x5
         * -4117 = xlHAlignDistributed           ->
         * 1     = xlHAlignGeneral               -> ALIGN_GENERAL = 0x0
         * 5     = xlHAlignFill                  -> ALIGN_FILL = 0x4
         * 7     = xlHAlignCenterAcrossSelection ->
         */
        private int getFeatureAlignStyle(int align)
        {
            switch (align)
            {
                case -4131: return 1;
                case -4108: return 2;
                case -4152: return 3;
                case -4130: return 5;
                case 5: return 4;
            }
            return 0;
        }

        private string getFeatureBorderStyle(Borders border)
        {
            string ret = "";
            // XlBordersIndex XlLineStyle
            ret += (border.get_Item(XlBordersIndex.xlEdgeTop).LineStyle != (int)XlLineStyle.xlLineStyleNone ? "1" : "0");
            ret += (border.get_Item(XlBordersIndex.xlEdgeBottom).LineStyle != (int)XlLineStyle.xlLineStyleNone ? "1" : "0");
            ret += (border.get_Item(XlBordersIndex.xlEdgeLeft).LineStyle != (int)XlLineStyle.xlLineStyleNone ? "1" : "0");
            ret += (border.get_Item(XlBordersIndex.xlEdgeRight).LineStyle != (int)XlLineStyle.xlLineStyleNone ? "1" : "0");
            return ret;
        }

        private int getFeatureFontBold(Font font)
        {
            if (font.Bold.ToString().Length == 0)
            {
                return 0;
            }
            return (font.Bold ? 1 : 0);
        }

        private double getFeatureFontHeight(Font font)
        {
            return font.Size;
        }

        private int getFeatureFontItalic(Font font)
        {
            return (font.Italic ? 1 : 0);
        }

        private int getFeatureFontUnderline(Font font)
        {
            return (font.Underline != (int)XlUnderlineStyle.xlUnderlineStyleNone ? 1 : 0);
        }

        private DataType getDataType(string numberFormat, string cellValue)
        {
            if (this.isDate(numberFormat))
            {
                return DataType.Date;
            }
            else if (this.isNumber(numberFormat) || this.isScientific(numberFormat))
            {
                return DataType.Number;
            }
            else if (this.isZipCode(numberFormat))
            {
                return DataType.ZipCode;
            }
            else if (this.isPhoneNumber(numberFormat))
            {
                return DataType.PhoneNumber;
            }
            else if (this.isText(numberFormat))
            {
                return DataType.Text;
            }
            if (this.isGeneral(numberFormat))
            {
                string valueType = this.getValueType(cellValue);
                if ("int".Equals(valueType) || "double".Equals(valueType))
                {
                    return DataType.Number;
                }
            }
            return DataType.General;
        }

        private void findColumnType(DataType[,] typeTable, int rowNum, int colNum, DataType[] columnTypeTable)
        {
            for (int j = 0; j < colNum; ++j)
            {
                int dateCount = 0, numberCount = 0, zipCodeCount = 0, phoneCount = 0, textCount = 0, generalCount = 0;
                int maxCount = int.MinValue;
                DataType maxCountType = DataType.NONE;
                for (int i = 0; i < rowNum; ++i)
                {
                    switch (typeTable[i, j])
                    {
                        case DataType.Date: ++dateCount; break;
                        case DataType.Number: ++numberCount; break;
                        case DataType.ZipCode: ++zipCodeCount; break;
                        case DataType.PhoneNumber: ++phoneCount; break;
                        case DataType.Text: ++textCount; break;
                        case DataType.General: ++generalCount; break;
                    }
                }
                if (dateCount > maxCount)
                {
                    maxCount = dateCount;
                    maxCountType = DataType.Date;
                }
                if (numberCount > maxCount)
                {
                    maxCount = numberCount;
                    maxCountType = DataType.Number;
                }
                if (zipCodeCount > maxCount)
                {
                    maxCount = zipCodeCount;
                    maxCountType = DataType.ZipCode;
                }
                if (phoneCount > maxCount)
                {
                    maxCount = phoneCount;
                    maxCountType = DataType.PhoneNumber;
                }
                if (textCount > maxCount)
                {
                    maxCount = textCount;
                    maxCountType = DataType.Text;
                }
                if (generalCount > maxCount)
                {
                    maxCount = generalCount;
                    maxCountType = DataType.General;
                }
                columnTypeTable[j] = maxCountType;
            }
        }

        #region DataType
        private bool isGeneral(string type)
        {
            if (type == "General")
            {
                return true;
            }
            return false;
        }
        private bool isNumber(string type)
        {
            Regex regex = new Regex(@"(#,##)?0.0*");
            return regex.Match(type).Success;
        }
        private bool isDate(string type)
        {
            Regex regex1 = new Regex(@"m/dd?/yy+");
            Regex regex2 = new Regex(@"mmmm dd?, yyyy");
            Regex regex3 = new Regex(@"(dd?-)?mmm+-yy+");
            Regex regex4 = new Regex(@"(h:)?mm:ss");
            Regex regex5 = new Regex(@"[h]:mm:ss");
            return regex1.Match(type).Success || regex2.Match(type).Success || regex3.Match(type).Success
                || regex4.Match(type).Success || regex5.Match(type).Success;
        }
        private bool isZipCode(string type)
        {
            if (type == "00000" || type == "00000-0000")
            {
                return true;
            }
            return false;
        }
        private bool isSocialSecurityNumber(string type)
        {
            if (type == "000-00-0000")
            {
                return true;
            }
            return false;
        }
        private bool isScientific(string type)
        {
            Regex regex = new Regex(@"0.0*E+00");
            return regex.Match(type).Success;
        }
        private bool isText(string type)
        {
            if (type == "@")
            {
                return true;
            }
            return false;
        }
        private bool isPhoneNumber(string type)
        {
            Regex regex = new Regex(@"###-####");
            return regex.Match(type).Success;
        }
        #endregion
    }
}
