using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameFinder
{
    class PredictSheetRow
    {
        public FeatureSheetRow feaRow;
        public PredictSheetRow()
        {
            this.feaRow = new FeatureSheetRow();
        }

        public void init()
        {
            Console.WriteLine("clean temp folder");
            string rmCmd = "rm";
            Process pid = new Process();
            pid.StartInfo.FileName = rmCmd;
            pid.StartInfo.UseShellExecute = false;
            pid.StartInfo.CreateNoWindow = true;
            string parameters = Config.CRFTEMPDIR + "/*/*";
            pid.StartInfo.Arguments = parameters;
            try
            {
                pid.Start();
                pid.WaitForExit();
                pid.Dispose();
            }
            catch
            {
                Console.WriteLine("No command found: " + rmCmd);
            }
        }

        public IEnumerable<Tuple<string, string, MSheet>> ScanEachExcel()
        {
            int counter = 0;
            string[] files = Directory.GetFiles(Config.SHEETDIR);
            for (int i = 0; i < files.Length; ++i)
            {
                string fileName = Path.GetFileName(files[i]);
                if (!fileName.EndsWith("xlsx") || fileName.StartsWith("~$"))
                {
                    continue;
                }
                /*
                try
                {
                    Console.WriteLine("Processing " + fileName);
                    SheetLoader sheetLoader = new SheetLoader(Config.SHEETDIR + "/" + fileName);
                    foreach(Tuple<string, MSheet> sheetDict in sheetLoader.FetchSheetDict())
                    {
                        string sheetName = sheetDict.Item1;
                        MSheet mSheet = sheetDict.Item2;
                        this.ProcessTable(fileName, sheetName, mSheet);
                        yield return mSheet;
                    }
                    ++counter;
                    if (counter % 100 == 0)
                    {
                        Console.WriteLine("Current: " + counter);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error while processing " + fileName + " " + e.Message);
                }
                */
                Console.WriteLine("Processing " + fileName);
                SheetLoader sheetLoader = new SheetLoader(Config.SHEETDIR + "/" + fileName);
                foreach (Tuple<string, MSheet> sheetDict in sheetLoader.FetchSheetDict())
                {
                    string sheetName = sheetDict.Item1;
                    MSheet mSheet = sheetDict.Item2;
                    this.ProcessTable(fileName, sheetName, mSheet);
                    yield return new Tuple<string, string, MSheet>(fileName, sheetName, mSheet);
                }
                sheetLoader.CloseWorkbook();
                ++counter;
                if (counter % 100 == 0)
                {
                    Console.WriteLine("Current: " + counter);
                }
            }
        }

        public void ProcessTable(string fileName, string sheetName, MSheet mSheet)
        {
            Dictionary<int, List<bool>> feaDict = this.feaRow.GenerateSingularFeatureCRF(mSheet);
            string outPath = Config.CRFTMPFEATURE;
            StreamWriter fout = new StreamWriter(outPath);
            List<int> keySetFromFeaDict = new List<int>(feaDict.Keys);
            keySetFromFeaDict.Sort();
            foreach (int row in keySetFromFeaDict)
            {
                List<bool> feaVec = feaDict[row];
                int a = row - 1;
                fout.Write(fileName + "____" + sheetName.Replace(" ", "__") + "____" + a + " ");
                foreach (bool feature in feaVec)
                {
                    if (feature)
                    {
                        fout.Write("1 ");
                    }
                    else
                    {
                        fout.Write("0 ");
                    }
                }
                fout.WriteLine("Title");
            }
            fout.Close();
        }
        /*
        public void generateFromSheetFile(string fileName)
        {
            SheetLoader sheetLoad = new SheetLoader(Config.SHEETDIR + "/" + fileName);
            Dictionary<string, MSheet> sheetDict = sheetLoad.LoadSheetDict();
            // Dictionary<string, MSheet> sheetDict = sheetLoad.LoadSheetDictByTransposition();
            foreach (string sheetName in sheetDict.Keys)
            {
                MSheet mSheet = sheetDict[sheetName];
                Dictionary<int, List<bool>> feaDict = this.feaRow.GenerateSingularFeatureCRF(mSheet);
                string outPath = Config.CRFFEADIR + "/" + fileName + "____" + sheetName;
                StreamWriter fout = new StreamWriter(outPath);
                List<int> keySetFromFeaDict = new List<int>(feaDict.Keys);
                keySetFromFeaDict.Sort();
                foreach (int row in keySetFromFeaDict)
                {
                    List<bool> feaVec = feaDict[row]; 
                    int a = row - 1;
                    fout.Write(fileName + "____" + sheetName.Replace(" ", "__") + "____" + a + " ");
                    foreach (bool feature in feaVec)
                    {
                        if (feature)
                        {
                            fout.Write("1 ");
                        }
                        else
                        {
                            fout.Write("0 ");
                        }
                    }
                    fout.WriteLine("Title");
                }
                fout.Close();
            }
            sheetLoad.CloseWorkbook();
        }
        */
    }
}
