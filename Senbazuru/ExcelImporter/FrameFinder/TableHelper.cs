using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;

namespace FrameFinder
{
    class TableHelper
    {
        private static int[] dx = { +1, -1, 0, 0, +1, -1, +1, -1 };
        private static int[] dy = { 0, 0, +1, -1, -1, -1, +1, +1 };
        private static int step = 4;
        public static List<Tuple<int, int, int, int>> SplitTable(Worksheet sheet, int threshold = 2)
        {
            List<Tuple<int, int, int, int>> ret = new List<Tuple<int, int, int, int>>();
            int rowNum = sheet.UsedRange.Cells.Rows.Count;
            int colNum = sheet.UsedRange.Cells.Columns.Count;
            int stRow = sheet.UsedRange.Row;
            int stCol = sheet.UsedRange.Column;
            int rangeUpRow = stRow;
            int rangeLeftCol = stCol;
            int rangeDownRow = stRow + rowNum - 1;
            int rangeRightCol = stCol + colNum - 1;
            bool[,] vis = new bool[rowNum, colNum];
            int[,] counter = new int[rowNum, colNum];
            for (int i = 0; i < rowNum; ++i)
            {
                for (int j = 0; j < colNum; ++j)
                {
                    counter[i, j] = threshold;
                }
            }

            for (int rowIdx = rangeUpRow; rowIdx <= rangeDownRow; ++rowIdx)
            {
                for (int colIdx = rangeLeftCol; colIdx <= rangeRightCol; ++colIdx)
                {
                    if (vis[rowIdx - stRow, colIdx - stCol])
                    {
                        continue;
                    }
                    Range cell = sheet.Cells[rowIdx, colIdx];
                    string value = Convert.ToString(cell.Value2);
                    if (value == null || value.Length == 0)
                    {
                        continue;
                    }
                    Queue<Tuple<int, int>> q = new Queue<Tuple<int, int>>();
                    q.Enqueue(new Tuple<int, int>(rowIdx, colIdx));
                    vis[rowIdx - stRow, colIdx - stCol] = true;

                    int minRow = int.MaxValue;
                    int minCol = int.MaxValue;
                    int maxRow = int.MinValue;
                    int maxCol = int.MinValue;

                    while (q.Count() > 0)
                    {
                        Tuple<int, int> cellCordinate = q.Dequeue();
                        int row = cellCordinate.Item1;
                        int col = cellCordinate.Item2;
                        cell = sheet.Cells[row, col];
                        if (counter[row - stRow, col - stCol] == 0)
                        {
                            continue;
                        }

                        if (cell.MergeCells)
                        {
                            counter[row - stRow, col - stCol] = threshold;
                            int upRow = cell.MergeArea.Row;
                            int leftCol = cell.MergeArea.Column;
                            int rowCount = cell.MergeArea.Rows.Count;
                            int colCount = cell.MergeArea.Columns.Count;
                            for (int j = leftCol; j < leftCol + colCount; ++j)
                            {
                                if (!vis[upRow - stRow, j - stCol])
                                {
                                    q.Enqueue(new Tuple<int, int>(upRow, j));
                                }
                                if (!vis[upRow + rowCount - 1 - stRow, j - stCol])
                                {
                                    q.Enqueue(new Tuple<int, int>(upRow + rowCount - 1, j));
                                }
                            }
                            for (int i = upRow; i < upRow + rowCount; ++i)
                            {
                                if (!vis[i - stRow, leftCol - stCol])
                                {
                                    q.Enqueue(new Tuple<int, int>(i, leftCol));
                                }
                                if (!vis[i - stRow, leftCol + colCount - 1 - stCol])
                                {
                                    q.Enqueue(new Tuple<int, int>(i, leftCol + colCount - 1));
                                }
                            }
                            for (int i = upRow; i < upRow + rowCount; ++i)
                            {
                                for (int j = leftCol; j < leftCol + colCount; ++j)
                                {
                                    vis[i - stRow, j - stCol] = true;
                                    counter[i - stRow, j - stCol] = threshold;
                                }
                            }
                        }

                        minRow = Math.Min(row, minRow);
                        minCol = Math.Min(col, minCol);
                        maxRow = Math.Max(row, maxRow);
                        maxCol = Math.Max(col, maxCol);

                        for (int i = 0; i < step; ++i)
                        {
                            int nextRow = row + dx[i];
                            int nextCol = col + dy[i];
                            if (nextRow >= rangeUpRow && nextRow <= rangeDownRow
                                && nextCol >= rangeLeftCol && nextCol <= rangeRightCol
                                && !vis[nextRow - stRow, nextCol - stCol])
                            {
                                vis[nextRow - stRow, nextCol - stCol] = true;
                                string cellValue = Convert.ToString(sheet.Cells[nextRow, nextCol].Value2);
                                if (cellValue == null || cellValue.Length == 0)
                                {
                                    counter[nextRow - stRow, nextCol - stCol] = counter[row - stRow, col - stCol] - 1;
                                }
                                q.Enqueue(new Tuple<int, int>(nextRow, nextCol));
                            }
                        }
                    }
                    if (minRow == int.MaxValue || minCol == int.MaxValue
                        || maxRow == int.MinValue || maxCol == int.MinValue)
                    {
                        continue;
                    }
                    ret.Add(new Tuple<int, int, int, int>(minRow, minCol, maxRow, maxCol));
                }
            }
            return TableHelper.Trim(sheet, ret);
        }

        public static List<Tuple<int, int, int, int>> Trim(Worksheet sheet, List<Tuple<int, int, int, int>> list)
        {
            List<Tuple<int, int, int, int>> ret = new List<Tuple<int,int,int,int>>();
            foreach (Tuple<int, int, int, int> cellRange in list)
            {
                int upRow = cellRange.Item1;
                int leftCol = cellRange.Item2;
                int downRow = cellRange.Item3;
                int rightCol = cellRange.Item4;
                
                for (int i = upRow; i <= downRow; ++i)
                {
                    bool isEmpty = true;
                    for (int j = leftCol; j <= rightCol; ++j)
                    {
                        Range cell = sheet.Cells[i, j];
                        string cellValue = Convert.ToString(cell.Value2);
                        if (cellValue == null)
                        {
                            continue;
                        }
                        if (cellValue.Trim().Length != 0)
                        {
                            isEmpty = false;
                            break;
                        }
                    }
                    if (isEmpty)
                    {
                        ++upRow;
                    }
                    else
                    {
                        break;
                    }
                }

                for (int i = downRow; i >= upRow; --i)
                {
                    bool isEmpty = true;
                    for (int j = leftCol; j <= rightCol; ++j)
                    {
                        Range cell = sheet.Cells[i, j];
                        string cellValue = Convert.ToString(cell.Value2);
                        if (cellValue == null)
                        {
                            continue;
                        }
                        if (cellValue.Trim().Length != 0)
                        {
                            isEmpty = false;
                            break;
                        }
                    }
                    if (isEmpty)
                    {
                        --downRow;
                    }
                    else
                    {
                        break;
                    }
                }

                for (int j = leftCol; j <= rightCol; ++j)
                {
                    bool isEmpty = true;
                    for (int i = upRow; i <= downRow; ++i)
                    {
                        Range cell = sheet.Cells[i, j];
                        string cellValue = Convert.ToString(cell.Value2);
                        if (cellValue == null)
                        {
                            continue;
                        }
                        if (cellValue.Trim().Length != 0)
                        {
                            isEmpty = false;
                            break;
                        }
                    }
                    if (isEmpty)
                    {
                        ++leftCol;
                    }
                    else
                    {
                        break;
                    }
                }

                for (int j = rightCol; j >= leftCol; --j)
                {
                    bool isEmpty = true;
                    for (int i = upRow; i <= downRow; ++i)
                    {
                        Range cell = sheet.Cells[i, j];
                        string cellValue = Convert.ToString(cell.Value2);
                        if (cellValue == null)
                        {
                            continue;
                        }
                        if (cellValue.Trim().Length != 0)
                        {
                            isEmpty = false;
                            break;
                        }
                    }
                    if (isEmpty)
                    {
                        --rightCol;
                    }
                    else
                    {
                        break;
                    }
                }

                ret.Add(new Tuple<int, int, int, int>(upRow, leftCol, downRow, rightCol));
            }
            return ret;
        }

        private static void print(List<Tuple<int, int, int, int>> list)
        {
            foreach (Tuple<int, int, int, int> tuple in list)
            {
                Console.WriteLine("{0}, {1}, {2}, {3}", tuple.Item1, tuple.Item2, tuple.Item3, tuple.Item4);
            }
        }
    }
}
