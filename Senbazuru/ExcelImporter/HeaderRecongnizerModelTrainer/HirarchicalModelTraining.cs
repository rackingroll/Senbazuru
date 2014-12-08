using Microsoft.Office.Interop.Excel;
using Senbazuru.HirarchicalExtraction;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Senbazuru.HirarchicalExtraction
{
    public class HirarchicalModel
    {
        static string Filename = @"..\..\..\resources\HirarchicalExtraction\TrainingData";
        Application excelapp = new Application();

        // We obtain the 1-th sheet
        private int SHEETNUM = 1;

        private string REGULAREXPRESSION_FILE = "0*.xls*";
        private string REGULAREXPRESSION_LABEL = "0*.label";

        List<AnotationPair> anotationPairTrainingList = new List<AnotationPair>();
        List<AnotationPairEdge> anotationPairEdgeTrainingList = new List<AnotationPairEdge>();

        List<AnotationPair> anotationPairTestingList = new List<AnotationPair>();
        List<AnotationPairEdge> anotationPairEdgeTestingList = new List<AnotationPairEdge>();

        List<AnotationPair> anotationPairGroundTruthList = new List<AnotationPair>();
        List<AnotationPairEdge> anotationPairEdgeGroundTruthList = new List<AnotationPairEdge>();


        private double TRAININGSIZE = 0.75;

        public HirarchicalModel(){}

        private AutomaticExtractionModel model = new AutomaticExtractionModel();

        /// <summary>
        /// Loading the training data from the file
        /// </summary>
        /// <param name="isTest">File loading</param>
        public void HirarchicalModelFileLoading(bool isTest)
        {
            System.IO.DirectoryInfo Dir = new System.IO.DirectoryInfo(Filename);

            FileInfo[] files = Dir.GetFiles(REGULAREXPRESSION_FILE);
            FileInfo[] labels = Dir.GetFiles(REGULAREXPRESSION_LABEL);

            if (files.Length != labels.Length) return;

            int trainingFileNum = (int)(files.Length * this.TRAININGSIZE);

            // Loading the training data
            for (int i = 0; i < trainingFileNum; i++)
            {
                FileInfo file = files[i];
                FileInfo label = labels[i];

                this.anotationPairTrainingList.AddRange(LabelReading(file.FullName, label.FullName,true,false));
                Console.WriteLine("Training File " + i + "Loaded");
            }

            if (!isTest) return;

            // Loading the testing data and ground truth data
            for (int i = trainingFileNum + 10; i < trainingFileNum + 11; i++)
            {
                FileInfo file = files[i];
                FileInfo label = labels[i];

                this.anotationPairGroundTruthList = LabelReading(file.FullName, label.FullName,true,false);
                this.anotationPairTestingList = LabelReading(file.FullName, label.FullName,false,true);
                Console.WriteLine("Testing File " + i + "Loaded");
            }
        }

        public void Train()
        {
            this.model.Training(this.anotationPairTrainingList, anotationPairEdgeTrainingList);
        }

        public void Testing()
        {
            this.model.Testing(this.anotationPairTestingList, this.anotationPairEdgeTestingList);
        }

        public void Evaluation()
        {
            double accuracy = 0.0;
            
            for (int i = 0; i < anotationPairTestingList.Count; i++)
            {
                int index = this.Indexof(anotationPairGroundTruthList, anotationPairTestingList[i]);
                if ( index != -1)
                {
                    if (anotationPairTestingList[i].nodepotentialfeaturevector.label == anotationPairGroundTruthList[index].nodepotentialfeaturevector.label)
                    {
                        accuracy++;
                    }
                }
                else
                {
                    if (anotationPairTestingList[i].nodepotentialfeaturevector.label == false)
                    {
                        accuracy++;
                    }
                }
            }

            accuracy = accuracy / anotationPairTestingList.Count;
            Console.WriteLine("The accuracy is: " + accuracy);
        }

        /// <summary>
        /// Reading information from the file
        /// </summary>
        /// <param name="file"></param>
        /// <param name="label"></param>
        /// <param name="addLabel">true: to add label information in the test</param>
        /// <param name="randomSample">
        /// true: random select some pairs in the area; 
        /// false: select the labeled pairs in the area.
        /// </param>
        /// <returns></returns>
        private List<AnotationPair> LabelReading (string file, string label, bool addLabel, bool randomSample)
        {
            Workbook workbook = excelapp.Workbooks.Open(file);
            Worksheet sheet = workbook.Sheets[SHEETNUM];

            StreamReader labelReader = new StreamReader(label);
            String value = labelReader.ReadLine();
            int startIdx = int.Parse(value.Split(' ')[1]);
            Range startCell = sheet.Cells[int.Parse(value.Split(' ')[1]), int.Parse(value.Split(' ')[2])];
            Range endCell = sheet.Cells[int.Parse(value.Split(' ')[3]), int.Parse(value.Split(' ')[4])];
            Range range = sheet.get_Range(startCell.get_Address() + ":" + endCell.get_Address());

            List<Tuple<int,int>> indexPairList = new List<Tuple<int,int>>() ;
            value = labelReader.ReadLine() ;
            while (value != null)
            {
                Tuple<int, int> indexPair = new Tuple<int, int>(int.Parse(value.Split(',')[0]) - startIdx, int.Parse(value.Split(',')[1]) - startIdx);
                indexPairList.Add(indexPair);
                value = labelReader.ReadLine();
            }

            
            FeatureConstructer constructer;
            if (!randomSample)
            {
                constructer = new FeatureConstructer(sheet, range, indexPairList);
            }
            else
            {
                constructer = new FeatureConstructer(sheet, range);
            }
            List<AnotationPair> anotationPairList = constructer.anotationPairList ;

            if (addLabel && !randomSample)
            {
                for (int i = 0; i < anotationPairList.Count; i++)
                {
                    anotationPairList[i].nodepotentialfeaturevector.label = true;
                }
            }

            return anotationPairList;
        }

        public AutomaticExtractionModel GetModel()
        {
            return this.model;
        }

        public void SaveModel ()
        {
            this.model.SaveModel();
        }

        public void LoadModel()
        {
            this.model.LoadModel();
        }
        /* Below are tool methods*/

        /// <summary>
        /// Find the index of pair in list
        /// return -1 denotes list does not contains pair
        /// </summary>
        /// <param name="list"></param>
        /// <param name="pair"></param>
        /// <returns></returns>
        private int Indexof(List<AnotationPair> list, AnotationPair pair)
        {
            for (int i = 0; i < list.Count; i++ )
            {
                if (list[i].indexChild == pair.indexChild && list[i].indexParent == pair.indexParent)
                {
                    return i;
                }
            }
            return -1;
        }
    }
}
