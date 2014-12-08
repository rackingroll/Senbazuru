using Senbazuru.HirarchicalExtraction;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HeaderRecongnizerModelTrainer
{
    class Program
    {
        static void Main(string[] args)
        {
            // Model Training
            HirarchicalModel model = new HirarchicalModel();

            model.LoadModel();

            model.HirarchicalModelFileLoading(true);

            //model.Train();

            //model.SaveModel();

            model.Testing();

            model.Evaluation();
        }
    }
}
