using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FrameFinder
{
    class FeatureSheetRow
    {
        private static string[] naset = { "(na)", "n/a", "(n/a)", "(x)", "-", "--", "z", "..." };
        private static char[] spcharset = { '<', '#', '>', ';', '$' };
        public static HashSet<String> NASET = new HashSet<String>(naset);
        public static HashSet<char> SPCHARSET = new HashSet<char>(spcharset);

        public Dictionary<int, List<bool>> GenerateSingularFeatureCRF(MSheet mSheet)
        {
            Dictionary<int, List<bool>> feaDict = new Dictionary<int, List<bool>>();
            for (int i = mSheet.StartRow; i < mSheet.StartRow + mSheet.RowNum; ++i)
            {
                Dictionary<int, MCell> rowCellDict = new Dictionary<int, MCell>();
                for (int j = mSheet.StartCol; j < mSheet.StartCol + mSheet.ColNum; ++j)
                {
                    Tuple<int, int> tuple = new Tuple<int, int>(i, j);
                    if (mSheet.SheetDict.ContainsKey(tuple))
                    {
                        MCell mCell = mSheet.SheetDict[tuple];
                        rowCellDict.Add(j, mCell);
                    }
                }
                if (rowCellDict.Count() == 0)
                {
                    continue;
                }
                bool blankFlag = false;
                if (feaDict.ContainsKey(i - 1))
                {
                    blankFlag = false;
                }
                else
                {
                    blankFlag = true;
                }
                feaDict.Add(i, this.generateFeatureByRowCRF(i, rowCellDict, mSheet, blankFlag));
            }
            return feaDict;
        }

        private List<bool> generateFeatureByRowCRF(int crow, Dictionary<int, MCell> rowCellDict,
                MSheet mSheet, bool blankFlag)
        {
            List<bool> feavec = new List<bool>();
            String cLineText = "";
            foreach (int ccol in rowCellDict.Keys)
            {
                MCell mCell = rowCellDict[ccol];
                cLineText += mCell.Value + " ";
            }
            // layout feature
            feavec.Add(blankFlag);
            feavec.Add(this.featureHasMergedCell(crow, mSheet));
            feavec.Add(this.featureReachRightBound(rowCellDict, mSheet.MaxColNum));
            feavec.Add(this.featureReachLeftBound(rowCellDict));
            feavec.Add(this.featureIsOneColumn(rowCellDict));
            feavec.Add(this.featureHasCenterAlignCell(rowCellDict));
            feavec.Add(this.featureHasLeftAlignCell(rowCellDict));
            feavec.Add(this.featureHasBoldFontCell(rowCellDict));
            feavec.Add(this.featureIndentation(cLineText));

            // textual feature
            feavec.Add(this.featureStartWithTable(cLineText));
            feavec.Add(this.featureStartWithPunctation(cLineText));
            feavec.Add(this.featureNumberPercentHigh(rowCellDict));
            feavec.Add(this.featureDigitalPercentHigh(rowCellDict));
            feavec.Add(this.featureAlphabetaAllCapital(cLineText));
            feavec.Add(this.featureAlphabetaStartWithCapital(rowCellDict));
            feavec.Add(this.featureAlphabetaStartWithLowercase(rowCellDict));
            feavec.Add(this.featureAlphabetaCellnumPercentHigh(rowCellDict));
            feavec.Add(this.featureAlphabetaPercentHigh(cLineText));
            feavec.Add(this.featureContainSpecialChar(cLineText));
            feavec.Add(this.featureContainColon(cLineText)); //
            feavec.Add(this.featureYearRangeCellnumHigh(rowCellDict));
            feavec.Add(this.featureYearRangePercentHigh(rowCellDict));
            feavec.Add(this.featureWordLengthHigh(rowCellDict));
            return feavec;
        }

        private bool featureHasMergedCell(int crow, MSheet mSheet)
        {
            return mSheet.MergeRowSet.Contains(crow);
        }

        private bool featureReachRightBound(Dictionary<int, MCell> rowCellDict, int maxCol)
        {
            return rowCellDict.ContainsKey(maxCol);
        }
        private bool featureReachLeftBound(Dictionary<int, MCell> rowCellDict)
        {
            return rowCellDict.ContainsKey(1);
        }
        private bool featureIsOneColumn(Dictionary<int, MCell> rowCellDict)
        {
            return rowCellDict.Count() == 1;
        }
        private bool featureHasCenterAlignCell(Dictionary<int, MCell> rowCellDict)
        {
            foreach (int col in rowCellDict.Keys)
            {
                if (rowCellDict[col].CenterAlignFlag)
                {
                    return true;
                }
            }
            return false;
        }
        private bool featureHasLeftAlignCell(Dictionary<int, MCell> rowCellDict)
        {
            foreach (int col in rowCellDict.Keys)
            {
                if (rowCellDict[col].LeftAlignFlag)
                {
                    return true;
                }
            }
            return false;
        }
        private bool featureHasBoldFontCell(Dictionary<int, MCell> rowCellDict)
        {
            foreach (int col in rowCellDict.Keys)
            {
                if (rowCellDict[col].BoldFlag)
                {
                    return true;
                }
            }
            return false;
        }
        private bool featureIndentation(String text)
        {
            int i = 0;
            for (; i < text.Length; ++i)
            {
                char c = text[i];
                if (c >= 'A' && c <= 'Z')
                {
                    break;
                }
                if (c >= 'a' && c <= 'z')
                {
                    break;
                }
                if (c >= '0' && c <= '9')
                {
                    break;
                }
            }
            return i > 0;
        }
        private bool featureStartWithTable(String text)
        {
            return text.Length > 0 && text.Trim().StartsWith("Table");
        }
        private bool featureStartWithPunctation(String text)
        {
            return text.Length > 0 && !this.hasDigits(text.Substring(0, 1))
                    && !this.hasLetter(text.Substring(0, 1));
        }
        private bool featureNumberPercentHigh(Dictionary<int, MCell> rowCellDict)
        {
            if (rowCellDict.Count() == 0)
            {
                return false;
            }
            int digitalCount = 0;
            foreach (int col in rowCellDict.Keys)
            {
                MCell mCell = rowCellDict[col];
                if (this.hasDigits(mCell.Value))
                {
                    ++digitalCount;
                }
                else if (this.isNa(mCell.Value))
                {
                    ++digitalCount;
                }
            }
            if ((double)digitalCount / (double)rowCellDict.Count() >= 0.6)
            {
                return true;
            }
            return false;
        }
        private bool featureDigitalPercentHigh(Dictionary<int, MCell> rowCellDict)
        {
            if (rowCellDict.Count() == 0)
            {
                return false;
            }
            int numberCount = 0;
            foreach (int col in rowCellDict.Keys)
            {
                MCell mCell = rowCellDict[col];
                if (this.isNumber(mCell.Value))
                {
                    ++numberCount;
                }
                else if (this.isNa(mCell.Value))
                {
                    ++numberCount;
                }
            }
            if ((double)numberCount / (double)rowCellDict.Count() >= 0.6)
            {
                return true;
            }
            return false;
        }
        private bool featureAlphabetaAllCapital(String text)
        {
            bool flag = false;
            for (int i = 0; i < text.Length; ++i)
            {
                char c = text[i];
                if (c >= 'A' && c <= 'Z')
                {
                    flag = true;
                }
                else if (c >= 'a' && c <= 'z')
                {
                    flag = false;
                    break;
                }
            }
            return flag;
        }
        private bool featureAlphabetaStartWithCapital(Dictionary<int, MCell> rowCellDict)
        {
            foreach (int col in rowCellDict.Keys)
            {
                MCell mCell = rowCellDict[col];
                if (!mCell.Type.Equals("str"))
                {
                    continue;
                }
                if (mCell.Value.Length == 0)
                {
                    continue;
                }
                if (this.hasLetter(mCell.Value) && !(mCell.Value[0] >= 'A' && mCell.Value[0] <= 'Z'))
                {
                    return false;
                }
            }
            return true;
        }
        private bool featureAlphabetaStartWithLowercase(Dictionary<int, MCell> rowCellDict)
        {
            int col = int.MaxValue;
            foreach (int key in rowCellDict.Keys)
            {
                col = Math.Min(col, key);
            }
            MCell mCell = rowCellDict[col];
            if (mCell.Value.Length == 0)
            {
                return false;
            }
            char c = mCell.Value[0];
            if (this.hasLetter(mCell.Value) && (c >= 'a' && c <= 'z'))
            {
                return true;
            }
            return false;
        }
        private bool featureAlphabetaCellnumPercentHigh(Dictionary<int, MCell> rowCellDict)
        {
            int counter = 0;
            Regex regex = new Regex(@"[A-Za-z]");
            foreach (int col in rowCellDict.Keys)
            {
                MCell mCell = rowCellDict[col];
                if (!mCell.Type.Equals("str"))
                {
                    continue;
                }
                if (regex.Match(mCell.Value).Success)
                {
                    ++counter;
                }
            }
            return (double)counter / (double)rowCellDict.Count() >= 0.6;
        }
        private bool featureAlphabetaPercentHigh(String text)
        {
            int counter = 0;
            for (int i = 0; i < text.Length; ++i)
            {
                char c = text[i];
                if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
                {
                    ++counter;
                }
            }
            double ret = (double)counter / (double)text.Length;
            return ret >= 0.6;
        }
        private bool featureContainSpecialChar(String text)
        {
            for (int i = 0; i < text.Length; ++i)
            {
                if (SPCHARSET.Contains(text[i]))
                {
                    return true;
                }
            }
            return false;
        }
        private bool featureContainColon(String text)
        {
            Regex regex = new Regex(@":");
            return regex.Match(text).Success;
        }
        private bool featureYearRangeCellnumHigh(Dictionary<int, MCell> rowCellDict)
        {
            if (rowCellDict.Count() == 0)
            {
                return false;
            }
            int yearCount = 0;
            foreach (int col in rowCellDict.Keys)
            {
                MCell mCell = rowCellDict[col];
                List<double> numArr = this.getNumberSet(mCell.Value);
                foreach (double year in numArr)
                {
                    if (year >= 1800 && year <= 2300)
                    {
                        ++yearCount;
                    }
                }
            }
            return yearCount >= 3;
        }
        private bool featureYearRangePercentHigh(Dictionary<int, MCell> rowCellDict)
        {
            if (rowCellDict.Count() == 0)
            {
                return false;
            }
            int yearCount = 0;
            // TODO total should be 1 or 0?
            int total = 1;
            foreach (int col in rowCellDict.Keys)
            {
                MCell mCell = rowCellDict[col];
                List<double> numArr = this.getNumberSet(mCell.Value);
                total += numArr.Count();
                foreach (double year in numArr)
                {
                    if (year >= 1800 && year <= 2300)
                    {
                        ++yearCount;
                    }
                }
            }
            return (double)yearCount / (double)total >= 0.7;
        }
        private bool featureWordLengthHigh(Dictionary<int, MCell> rowCellDict)
        {
            if (rowCellDict.Count() != 1)
            {
                return false;
            }
            foreach (int col in rowCellDict.Keys)
            {
                MCell mCell = rowCellDict[col];
                if (mCell.Value.Length > 40)
                {
                    return true;
                }
            }
            return false;
        }

        private bool hasLetter(String s)
        {
            for (int i = 0; i < s.Length; ++i)
            {
                char c = s[i];
                if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
                {
                    return true;
                }
            }
            return false;
        }

        private bool hasDigits(String s)
        {
            for (int i = 0; i < s.Length; ++i)
            {
                char c = s[i];
                if (c >= '0' && c <= '9')
                {
                    return true;
                }
            }
            return false;
        }
        private bool isNa(String s)
        {
            return NASET.Contains(s.Trim().ToLower());
        }
        private bool isNumber(String s)
        {
            try
            {
                Double.Parse(s);
            }
            catch
            {
                return false;
            }
            return true;
        }
        private List<double> getNumberSet(String s)
        {
            // TODO should check int or double
            String[] arr = s.Split(" ".ToCharArray());
            List<double> ret = new List<double>();
            for (int i = 0; i < arr.Length; ++i)
            {
                double t;
                try
                {
                    t = Double.Parse(arr[i]);
                }
                catch
                {
                    continue;
                }
                ret.Add(t);
            }
            return ret;
        }

        public class FeatureFormat
        {
            FeatureFormat() { }

        }
    }
}
