using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;

namespace FrameFinder
{
    class HorizontalHierarchyExtractor
    {

        public HorizontalHierarchyExtractor() { }
        public void ExtractFromMSheet(MSheet mSheet)
        {
            int rowNum = mSheet.RowNum;
            int colNum = mSheet.ColNum;
            this.initNodes(mSheet, rowNum, colNum);
            bool[,] vis = new bool[rowNum, colNum];
            for (int i = 0; i < rowNum; ++i)
            {
                for (int j = 0; j < colNum; ++j)
                {
                    vis[i, j] = false;
                }
            }
            int upRow = mSheet.StartRow;
            int leftCol = mSheet.StartCol;
            int downRow = mSheet.StartRow + mSheet.RowNum - 1;
            int rightCol = mSheet.StartCol + mSheet.ColNum - 1;
            
            for (int j = leftCol; j <= rightCol; ++j)
            {
                for (int i = upRow; i <= downRow; ++i)
                {
                    if (vis[i - upRow, j - leftCol])
                    {
                        continue;
                    }
                    Range cell = mSheet.Cells.Cells[i, j];
                    int row = i;
                    int col = j;
                    if (cell.MergeCells)
                    {
                        row = cell.MergeArea.Row;
                        col = cell.MergeArea.Column;
                        for (int ii = 0; ii < cell.MergeArea.Rows.Count; ++ii)
                        {
                            for (int jj = 0; jj < cell.MergeArea.Columns.Count; ++jj)
                            {
                                vis[row + ii - mSheet.StartRow, col + jj - mSheet.StartCol] = true;
                            }
                        }
                    }
                    vis[i - upRow, j - leftCol] = true;
                    Tuple<int, int> tmp = new Tuple<int, int>(row, col);
                    if (mSheet.Nodes.Keys.Contains(tmp))
                    {
                        int preRow = row - 1;
                        int preCol = col;
                        Range preCell = mSheet.Cells.Cells[preRow, preCol];
                        if (preCell.MergeCells)
                        {
                            preRow = preCell.MergeArea.Row;
                            preCol = preCell.MergeArea.Column;
                        }
                        Tuple<int, int> preTmp = new Tuple<int, int>(preRow, preCol);
                        if (mSheet.Nodes.Keys.Contains(preTmp))
                        {
                            mSheet.Nodes[preTmp].AddChild(mSheet.Nodes[tmp]);
                            mSheet.Nodes[tmp].Parent = mSheet.Nodes[preTmp];
                        }
                    }
                }
            }
            this.linkToRoot(mSheet);
        }
        private void initNodes(MSheet mSheet, int rowNum, int colNum)
        {
            mSheet.InitNodes();
            bool[,] vis = new bool[rowNum, colNum];
            for (int i = 0; i < rowNum; ++i)
            {
                for (int j = 0; j < colNum; ++j)
                {
                    vis[i, j] = false;
                }
            }
            for (int j = mSheet.StartCol; j < mSheet.StartCol + mSheet.ColNum; ++j)
            {
                for (int i = mSheet.StartRow; i < mSheet.StartRow + mSheet.RowNum; ++i)
                {
                    if (vis[i - mSheet.StartRow, j - mSheet.StartCol])
                    {
                        continue;
                    }
                    if (mSheet.Labels.Keys.Contains(i) && mSheet.Labels[i] == RowLabel.Header)
                    {
                        Range cell = mSheet.Cells.Cells[i, j];
                        int row = i;
                        int col = j;
                        if (cell.MergeCells)
                        {
                            row = cell.MergeArea.Row;
                            col = cell.MergeArea.Column;
                            for (int ii = 0; ii < cell.MergeArea.Rows.Count; ++ii)
                            {
                                for (int jj = 0; jj < cell.MergeArea.Columns.Count; ++jj)
                                {
                                    vis[row + ii - mSheet.StartRow, col + jj - mSheet.StartCol] = true;
                                }
                            }
                        }
                        vis[row - mSheet.StartRow, col - mSheet.StartCol] = true;
                        mSheet.Nodes.Add(new Tuple<int, int>(row, col), new HeaderNode(row, col));
                    }
                }
            }
        }

        private void linkToRoot(MSheet mSheet)
        {
            foreach (Tuple<int, int> key in mSheet.Nodes.Keys)
            {
                HeaderNode node = mSheet.Nodes[key];
                if (!node.HasParent())
                {
                    mSheet.RootNode.AddChild(node);
                    mSheet.Nodes[key].Parent = mSheet.RootNode;
                }
            }
            // this.printStructure(mSheet.RootNode, "");
        }

        private void printStructure(HeaderNode node, string space)
        {
            Console.WriteLine("{0} {1} {2}", space, node.Row, node.Col);
            foreach (HeaderNode child in node.Children)
            {
                this.printStructure(child, space + " ");
            }
        }
    }
}
