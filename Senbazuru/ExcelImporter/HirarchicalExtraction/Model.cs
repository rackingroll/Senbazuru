using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Senbazuru.HirarchicalExtraction
{
    public class Model
    {

        public int NUM_EDGE_POTENTIAL_FEATURE_NUMBER = 3;
        public int NUM_NODE_POTENTIAL_FEATURE_NUMBER = 14;

        // Parameters
        private double LAMBDA = 0.1;
        private double THRESHOLD = 0.001;
        private double LABELTHRESHOLD = 0.6;
        private double ADJUSTTHRESHOLD = 0.1;
        private double ADJUSTWEIGHT = 0.5;

        public IList<double> WeightListEdgeFeature = new List<double>();
        public IList<double> WeightListNodeFeature = new List<double>();

        public Model() {}
        
        /// <summary>
        /// Using the feature vector List and the weightList to obtain the label of each anotation pair
        /// </summary>
        /// <param name="featurevector">The feature vector without labeling</param>
        public void Testing(IList<NodePotentialFeatureVector> nodepotentialfeaturevector, IList<EdgePotentialFeatureVector> edgepotentialfeaturevector)
        {
            if (this.WeightListEdgeFeature.Count == 0 || this.WeightListNodeFeature.Count == 0)
            {
                return;
            }

            double label = 0.0 ;
            for (int i = 0; i < nodepotentialfeaturevector.Count; i++)
            {
                label = 0.0;
                for (int j = 0; j < this.NUM_NODE_POTENTIAL_FEATURE_NUMBER; j++)
                {
                    label += nodepotentialfeaturevector[i].features[j] * this.WeightListNodeFeature[j];
                }
                if (label > this.LABELTHRESHOLD)
                {
                    nodepotentialfeaturevector[i].label = true;
                }
                else
                {
                    nodepotentialfeaturevector[i].label = false;
                }
            }

            for (int i = 0; i < edgepotentialfeaturevector.Count; i++)
            {
                if (nodepotentialfeaturevector[edgepotentialfeaturevector[i].pair1Idx].label != nodepotentialfeaturevector[edgepotentialfeaturevector[i].pair2Idx].label)
                {
                    break;
                }
                label = 0.0;
                for (int j = 0; j < this.NUM_EDGE_POTENTIAL_FEATURE_NUMBER; j++)
                {
                    label += edgepotentialfeaturevector[i].features[j] * this.WeightListEdgeFeature[j];
                }
                if (label > this.LABELTHRESHOLD)
                {
                    edgepotentialfeaturevector[i].HasEdge = true;
                }
                else
                {
                    edgepotentialfeaturevector[i].HasEdge = false;
                }
            }
        }

        /// <summary>
        /// Using Gradient Descend
        /// A list of Feature vector with labels
        /// Obtain the WeightList
        /// </summary>
        /// <param name="featurevector"></param>
        public void Training(IList<NodePotentialFeatureVector> nodepotentialfeaturevector, IList<EdgePotentialFeatureVector> edgepotentialfeaturevector)
        {

            // initialization
            double Descend = 0.0; // the direction of descend
            for (int i = 0; i < this.NUM_EDGE_POTENTIAL_FEATURE_NUMBER; i++)
            {
                WeightListEdgeFeature.Add(0.0);
            }
            for (int i = 0; i < this.NUM_NODE_POTENTIAL_FEATURE_NUMBER; i++)
            {
                WeightListNodeFeature.Add(0.0);
            }

            while (true)
            {
                IList<double> WeightListOld = this.JoinWeightList() ;
                for (int k = 0; k < this.NUM_NODE_POTENTIAL_FEATURE_NUMBER; k++)
                {
                    Descend = NodePotentialDescend(nodepotentialfeaturevector,edgepotentialfeaturevector, k);
                    WeightListNodeFeature[k] += this.LAMBDA * Descend;
                }
                for (int k = 0; k < this.NUM_EDGE_POTENTIAL_FEATURE_NUMBER; k++)
                {
                    Descend = EdgePotentialDescend(nodepotentialfeaturevector, edgepotentialfeaturevector, k);
                    WeightListEdgeFeature[k] += this.LAMBDA * Descend;
                }

                this.FeatureNorm();
                if (this.Error(WeightListOld) < this.THRESHOLD)
                {
                    break;
                }
            }
            // Addjust feature, because weight of the feature need to have negative effect on the model, but this training process only considers the 
            //FeatureAdjust();
            
        }

        private double NodePotentialDescend(IList<NodePotentialFeatureVector> nodepotentialfeaturevector, IList<EdgePotentialFeatureVector> edgepotentialfeaturevector, int k)
        {
            double Descend = 0.0;
            for (int i = 0; i < nodepotentialfeaturevector.Count; i++)
            {
                Descend += nodepotentialfeaturevector[i].features[k];
            }
            return Descend - NormalizationTerm(nodepotentialfeaturevector, edgepotentialfeaturevector,k,3) - RegularizationParameter(1, k);
        }

        private double EdgePotentialDescend(IList<NodePotentialFeatureVector> nodepotentialfeaturevector, IList<EdgePotentialFeatureVector> edgepotentialfeaturevector, int k)
        {
            double Descend = 0.0;
            for (int i = 0; i < edgepotentialfeaturevector.Count; i++)
            {
                Descend += edgepotentialfeaturevector[i].features[k];
            }
            return Descend - NormalizationTerm(nodepotentialfeaturevector, edgepotentialfeaturevector, k,3) - RegularizationParameter(2, k);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="FeatureType">1. for node potential feature; 2. for edge potential feature</param>
        /// <param name="k"></param>
        /// <returns></returns>
        private double RegularizationParameter(int FeatureType, int k)
        {
            // Initial weightList
            IList<double> WeightList = this.JoinWeightList();

            double variance = this.Variance(WeightList) ;

            if (variance == 0) return 0.0;

            switch(FeatureType)
            {
                case 1:
                    return this.WeightListNodeFeature[k] / variance;
                case 2:
                    return this.WeightListEdgeFeature[k] / variance;
                default:
                    return 0.0;
            }
        }

        private double NormalizationTerm(IList<NodePotentialFeatureVector> nodepotentialfeaturevector, IList<EdgePotentialFeatureVector> edgepotentialfeaturevector, int k, int type)
        {
            double sum = 0.0 ;
            for (int i = 0; i < nodepotentialfeaturevector.Count; i++)
            {
                for (int j = 0; j < this.NUM_NODE_POTENTIAL_FEATURE_NUMBER; j++)
                {
                    sum += nodepotentialfeaturevector[i].features[j] * this.WeightListNodeFeature[j];
                }
            }
            for (int i = 0; i < edgepotentialfeaturevector.Count; i++)
            {
                for (int j = 0; j < this.NUM_EDGE_POTENTIAL_FEATURE_NUMBER; j++)
                {
                    sum += edgepotentialfeaturevector[i].features[j] * this.WeightListEdgeFeature[j];
                }
            }

            switch (type)
            {
                case 1:
                    return (Math.Exp(sum) * this.WeightListNodeFeature[k]) / (1.0 + Math.Exp(sum));
                case 2:
                    double a = (Math.Exp(sum) * this.WeightListEdgeFeature[k]) / (1.0 + Math.Exp(sum));
                    return (Math.Exp(sum) * this.WeightListEdgeFeature[k]) / (1.0 + Math.Exp(sum));
                default:
                    return 0.0;
            }
        }

        private double Variance(IList<double> list)
        {
            double variance = 0.0;
            double mean = list.Sum() / list.Count;
            for (int i = 0; i < list.Count; i++)
            {
                variance += (list[i] - mean) * (list[i] - mean);
            }
            return variance / list.Count;
        }

        private IList<double> JoinWeightList()
        {
            IList<double> WeightList = new List<double>();
            for (int i = 0; i < WeightListNodeFeature.Count; i++)
            {
                WeightList.Add(WeightListNodeFeature[i]);
            }
            for (int i = 0; i < WeightListEdgeFeature.Count; i++)
            {
                WeightList.Add(WeightListEdgeFeature[i]);
            }
            return WeightList;
        }

        private double Error(IList<double> weightlistold)
        {
            double error = 0.0;
            IList<double> WeightList = this.JoinWeightList() ;
            for (int i = 0; i < WeightList.Count; i++)
            {
                error += (weightlistold[i] - WeightList[i]) * (weightlistold[i] - WeightList[i]);
            }
            Console.WriteLine("Error = " + Math.Sqrt(error));
            return Math.Sqrt(error);
        }

        private void FeatureNorm()
        {
            
            double sum = this.WeightListNodeFeature.Sum();
            if (sum != 0.0)
            {
                for (int i = 0; i < this.WeightListNodeFeature.Count; i++)
                {
                    this.WeightListNodeFeature[i] /= sum;
                }
            }

            sum = this.WeightListEdgeFeature.Sum();
            if (sum != 0.0)
            {
                for (int i = 0; i < this.WeightListEdgeFeature.Count; i++)
                {
                    this.WeightListEdgeFeature[i] /= sum;
                }
            }
        }


        /*Below Method may be not used in our method!*/
        /// <summary>
        /// Addjust feature, because weight of the feature need to have negative effect on the model, but this training process only considers the 
        /// </summary>
        private void FeatureAdjust()
        {
            for (int i = 0; i < this.WeightListEdgeFeature.Count; i++)
            {
                if (this.WeightListEdgeFeature[i] < this.ADJUSTTHRESHOLD)
                {
                    this.WeightListEdgeFeature[i] -= this.ADJUSTWEIGHT;
                }
            }

            for (int i = 0; i < this.WeightListNodeFeature.Count; i++)
            {
                if (this.WeightListNodeFeature[i] < this.ADJUSTTHRESHOLD)
                {
                    this.WeightListNodeFeature[i] -= this.ADJUSTWEIGHT;
                }
            }
        }
    }
}
