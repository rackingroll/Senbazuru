using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameFinder
{
    class CRFRunner
    {
        string crfTrainScript;
        string crfTestScript;
        string crfModelPath;
        double cost;

        public CRFRunner()
        {
            this.crfTrainScript = Config.CRFPPDIR + "/crf_learn";
            this.crfTestScript = Config.CRFPPDIR + "/crf_test";
            this.crfModelPath = Config.CRFTEMPDIR + "/model";
            this.cost = 4.0;
        }

        public CRFRunner(double cost)
        {
            this.crfTrainScript = Config.CRFPPDIR + "/crf_learn";
            this.crfTestScript = Config.CRFPPDIR + "/crf_test";
            this.crfModelPath = Config.CRFTEMPDIR + "/model";
            this.cost = cost;
        }

        public void Train()
        {
            Console.WriteLine("Training CRF++ model...");
            Process pid = new Process();
            pid.StartInfo.FileName = this.crfTrainScript;
            pid.StartInfo.UseShellExecute = false;
            pid.StartInfo.CreateNoWindow = true;
            string parameters = string.Format(" -c {0} {1} {2} {3}", this.cost, Config.CRFPPTEMPLATEPATH,
                Config.CRFTRAINDATAPATH, this.crfModelPath);
            pid.StartInfo.Arguments = parameters;
            pid.StartInfo.RedirectStandardOutput = true;
            pid.Start();
            StreamReader sr = pid.StandardOutput;
            string line = null;
            while ((line = sr.ReadLine()) != null)
            {
                Console.WriteLine(line);
            }
            pid.WaitForExit();
            pid.Dispose();
            Console.WriteLine("Done training CRF++ model...");
        }
        public void Predict(string workbookName = "", string sheetName = "")
        {
            string featureTmpPath = Config.CRFTMPFEATURE;
            string fileName = Path.GetFileName(featureTmpPath);
            if (sheetName.Length == 0 || workbookName.Length == 0)
            {
                Console.WriteLine("CRF++ predicting row labels for table");
            }
            else
            {
                Console.WriteLine("CRF++ predicting row labels for table in sheet \"{0}\" of workbook \"{1}\"", sheetName, workbookName);
            }
            string predictTmpPath = Config.CRFTMPPREDICT;
            StreamWriter sw = new StreamWriter(predictTmpPath);
            Process pid = new Process();
            pid.StartInfo.FileName = this.crfTestScript;
            pid.StartInfo.UseShellExecute = false;
            pid.StartInfo.CreateNoWindow = true;
            string parameters = string.Format(" -m {0} {1}", this.crfModelPath, featureTmpPath);
            pid.StartInfo.Arguments = parameters;
            pid.StartInfo.RedirectStandardOutput = true;
            pid.Start();

            using (StreamReader sr = pid.StandardOutput)
            {
                string result = sr.ReadToEnd();
                sw.Write(result);
            }

            sw.Close();
            pid.WaitForExit();
            pid.Dispose();
		}
    }
}
