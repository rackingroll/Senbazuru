using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Senbazuru.HirarchicalExtraction
{
    public class AutomaticExtractionModel
    {
        private Model model = new Model();

        AttributeTree tree = new AttributeTree() ;

        public AutomaticExtractionModel() { }

        private string MODEL_FILE = @"..\..\..\resources\HirarchicalExtraction\Model\Hirarchical.model";

        /// <summary>
        /// Training the model
        /// </summary>
        /// <param name="AnotationPairList"></param>
        /// <param name="AnotationPairEdgeList"></param>
        public void Training(IList<AnotationPair> AnotationPairList, IList<AnotationPairEdge> AnotationPairEdgeList)
        {
            IList<NodePotentialFeatureVector> nodepotentialfeaturevector = new List<NodePotentialFeatureVector> ();
            IList<EdgePotentialFeatureVector> edgepotentialfeaturevector = new List<EdgePotentialFeatureVector> ();

            for (int i=0 ;i<AnotationPairList.Count;i++)
            {
                // only extract true parent-child as the label information
                if (AnotationPairList[i].nodepotentialfeaturevector.label == true)
                {
                    nodepotentialfeaturevector.Add(AnotationPairList[i].nodepotentialfeaturevector);
                }
            }
            for (int i = 0; i < AnotationPairEdgeList.Count; i++)
            {
                // Only got the exist edges.
                if (AnotationPairEdgeList[i].edgepotentialfeaturevector.HasEdge == true)
                {
                    edgepotentialfeaturevector.Add(AnotationPairEdgeList[i].edgepotentialfeaturevector);
                }
            }

            this.model.Training(nodepotentialfeaturevector, edgepotentialfeaturevector);
        }

        /// <summary>
        /// Add label for each AnotationPair in the List
        /// </summary>
        /// <param name="AnotationPairList"></param>
        public void Testing(IList<AnotationPair> AnotationPairList, IList<AnotationPairEdge> AnotationPairEdgeList)
        {
            IList<NodePotentialFeatureVector> nodepotentialfeaturevector = new List<NodePotentialFeatureVector>();
            IList<EdgePotentialFeatureVector> edgepotentialfeaturevector = new List<EdgePotentialFeatureVector>();
            for (int i = 0; i < AnotationPairList.Count; i++)
            {
                nodepotentialfeaturevector.Add(AnotationPairList[i].nodepotentialfeaturevector);
            }
            for (int i = 0; i < AnotationPairEdgeList.Count; i++)
            {
                edgepotentialfeaturevector.Add(AnotationPairEdgeList[i].edgepotentialfeaturevector);
            }
            this.model.Testing(nodepotentialfeaturevector, edgepotentialfeaturevector);

            // update the label information of the current anotationpairlist
            for (int i = 0; i < AnotationPairList.Count; i++)
            {
                AnotationPairList[i].nodepotentialfeaturevector = nodepotentialfeaturevector[i];
            }
            for (int i = 0; i < AnotationPairEdgeList.Count; i++)
            {
                AnotationPairEdgeList[i].edgepotentialfeaturevector = edgepotentialfeaturevector[i];
            }
        }

        public AttributeTree GetTree(IList<AnotationPair> AttributePairList)
        {
            this.tree.AttributeRelationDictionary(AttributePairList);
            return this.tree;
        }

        public AttributeTree Repairing(IList<AnotationPair> AttributePairList, IList<AnotationPairEdge> AttributePairEdgeList, int OldParentRowNum, int TargetRowNum, int NewParentRowNum)
        {
            this.tree.Repair(AttributePairList, AttributePairEdgeList, OldParentRowNum, TargetRowNum, NewParentRowNum);
            this.tree.AttributeRelationDictionary(AttributePairList);
            return this.tree;
        }


        /*Save and Load Model from Model file*/
        public void SaveModel()
        {
            StreamWriter writer = new StreamWriter(this.MODEL_FILE);

            for (int i = 0; i < this.model.WeightListNodeFeature.Count; i++)
            {
                writer.WriteLine(this.model.WeightListNodeFeature[i]);
            }

            writer.Close();
        }

        public void LoadModel()
        {
            StreamReader reader = new StreamReader(this.MODEL_FILE);

            if (this.model.WeightListNodeFeature.Count != 0)
            {
                for (int i = 0; i < this.model.NUM_NODE_POTENTIAL_FEATURE_NUMBER; i++)
                {
                    string value = reader.ReadLine();
                    this.model.WeightListNodeFeature[i] = double.Parse(value);
                }
            }
            else
            {
                for (int i = 0; i < this.model.NUM_NODE_POTENTIAL_FEATURE_NUMBER; i++)
                {
                    string value = reader.ReadLine();
                    this.model.WeightListNodeFeature.Add(double.Parse(value));
                }
            }
            reader.Close();
        }
    }
}
