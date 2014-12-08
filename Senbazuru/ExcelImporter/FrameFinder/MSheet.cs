using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;

namespace FrameFinder
{
    class MSheet
    {
        private string txt;
        private int maxRowNum;
        private int maxColNum;
        private int rowNum;
        private int colNum;
        private int stRow;
        private int stCol;

        public string Text 
        { 
            get { return this.txt; }
            set { this.txt = value; }
        }
        public int MaxRowNum
        {
            get { return this.maxRowNum; }
            set { this.maxRowNum = value; }
        }
        public int MaxColNum
        {
            get { return this.maxColNum; }
            set { this.maxColNum = value; }
        }
        public int RowNum
        {
            get { return this.rowNum; }
            set { this.rowNum = value; }
        }
        public int ColNum
        {
            get { return this.colNum; }
            set { this.colNum = value; }
        }
        public int StartRow
        {
            get { return this.stRow; }
            set { this.stRow = value; }
        }
        public int StartCol
        {
            get { return this.stCol; }
            set { this.stCol = value; }
        }

        public HashSet<int> MergeRowSet;
        public List<Tuple<int, int>> MergeCellSet;
        public Dictionary<Tuple<int, int>, MCell> SheetDict;
        public Dictionary<int, RowLabel> Labels;
        public DataType[] ColumnTypes;
        public Range Cells;

        public HeaderNode RootNode;
        public Dictionary<Tuple<int, int>, HeaderNode> Nodes;

        public MSheet()
        {
            this.MergeRowSet = new HashSet<int>();
            this.MergeCellSet = new List<Tuple<int, int>>();
            this.txt = "";
            this.maxRowNum = 1;
            this.maxColNum = 1;
            this.rowNum = 0;
            this.colNum = 0;
            this.stRow = 1;
            this.stCol = 1;
            this.SheetDict = new Dictionary<Tuple<int, int>, MCell>();
            this.Labels = new Dictionary<int, RowLabel>();
        }
        public void AddMergeCell(int row1, int row2, int col1, int col2)
        {
            for (int i = row1; i <= row2; ++i)
            {
                this.MergeRowSet.Add(i);
                for (int j = col1; j <= col2; ++j)
                {
                    this.MergeCellSet.Add(new Tuple<int, int>(i, j));
                }
            }
        }
        public void InsertCell(int rowIdx, int colIdx, string cType, int indents,
            int alignStyle, string borderStyle, int bgColor, int boldFlag,
            int height, int italicFlag, int underlineFlag, string value)
        {
            MCell mCell = new MCell();
            mCell.Init(value, cType, indents, alignStyle, boldFlag, borderStyle, bgColor,
                    height, italicFlag, underlineFlag);
            this.SheetDict.Add(new Tuple<int, int>(rowIdx, colIdx), mCell);
            if (cType.Equals("str"))
            {
                this.txt += value + " ";
            }
            this.maxRowNum = Math.Max(this.maxRowNum, rowIdx);
            this.maxColNum = Math.Max(this.maxColNum, colIdx);
        }

        /*
         * Choose the type appears maximum to represent the type of column
         */
        public void SetColumnTypeTable(DataType[] columnTypes, int num)
        {
            this.ColumnTypes = new DataType[num];
            for (int j = 0; j < num; ++j)
            {
                this.ColumnTypes[j] = columnTypes[j];
            }
        }

        public void InitNodes()
        {
            this.RootNode = new HeaderNode();
            this.Nodes = new Dictionary<Tuple<int, int>, HeaderNode>();
        }
    }
}
