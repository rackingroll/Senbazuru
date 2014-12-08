using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FrameFinder
{
    class TransformOutput
    {
        public static void Run(string workbookName, string sheetName, MSheet mSheet, int counter = 0)
        {
            string predictPath = Config.CRFTMPPREDICT;
            string fileName = Path.GetFileName(predictPath);
            Console.WriteLine("Generating final output");
            String outPath = Config.OUTPUTDIR + "/" + workbookName + "____" + sheetName + "____" + counter;
            StreamReader fin = new StreamReader(predictPath);
            StreamWriter fout = new StreamWriter(outPath);
            String line;
            while ((line = fin.ReadLine()) != null)
            {
                String[] strArr = Regex.Split(line.Trim(), @"\s+");
                if (strArr.Length == 0)
                {
                    continue;
                }
                String cKey = strArr[0];
                if (cKey.Length == 0)
                {
                    continue;
                }
                String label = strArr[strArr.Length - 1];
                String[] strArr2 = Regex.Split(cKey.Trim(), @"____");
                int row = int.Parse(strArr2[strArr2.Length - 1]);
                fout.WriteLine((row + 1).ToString() + "\t" + label);
                mSheet.Labels.Add(row + 1, DataTypes.String2RowLabel(label));
            }
            fin.Close();
            fout.Close();
            Console.WriteLine("Successfully obtain prediction results");
        }
    }
}
