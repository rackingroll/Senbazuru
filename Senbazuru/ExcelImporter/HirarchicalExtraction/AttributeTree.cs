using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Senbazuru.HirarchicalExtraction
{
    public class AttributeTree
    {
        public IList<Range> CellList = null;

        public Dictionary<int, int> AttributeChildtoParent = new Dictionary<int, int>();

        public AttributeTree() { }

        /// <summary>
        /// Construct the tree structure of the attribute model
        /// There are two rules in this model:
        /// 1. One child only have one parent. (The most adjacent parent is the one need to be chose)
        /// 2. Parent's index > Child's index
        /// </summary>
        /// <param name="AttributePairList"></param>
        public void AttributeRelationDictionary(IList<AnotationPair> AttributePairList)
        {
            this.CellList = AttributePairList[0].CellList ;
            for (int i = 0; i < AttributePairList.Count; i++)
            {
                if (AttributePairList[i].nodepotentialfeaturevector.label == true)
                {
                    if (AttributeChildtoParent.ContainsKey(AttributePairList[i].indexChild))
                    {
                        // current parent is more adjacent
                        if (Math.Abs(AttributeChildtoParent[AttributePairList[i].indexChild] - AttributePairList[i].indexChild) > Math.Abs(AttributePairList[i].indexParent - AttributePairList[i].indexChild)
                            && AttributePairList[i].indexParent < AttributePairList[i].indexChild)
                        {
                            AttributeChildtoParent[AttributePairList[i].indexChild] = AttributePairList[i].indexParent;
                        }
                    }
                    else
                    {
                        AttributeChildtoParent.Add(AttributePairList[i].indexChild, AttributePairList[i].indexParent);
                    }
                }
            }
        }

        public void Repair(IList<AnotationPair> AttributePairList, IList<AnotationPairEdge> AttributePairEdgeList, int OldParentRowNum, int TargetRowNum, int NewParentRowNum)
        {
            int index = this.RepairSinglePair(AttributePairList,OldParentRowNum,TargetRowNum,NewParentRowNum) ;

            for (int i = 0; i < AttributePairEdgeList.Count; i++)
            {
                if (AttributePairList[index] == AttributePairEdgeList[i].pair1
                    && AttributePairEdgeList[i].edgepotentialfeaturevector.HasEdge == true
                    && AttributePairEdgeList[i].pair2.indexParent == OldParentRowNum)
                {
                    this.RepairSinglePair(AttributePairList, OldParentRowNum, AttributePairEdgeList[i].pair2.indexChild, NewParentRowNum);
                }
                if (AttributePairList[index] == AttributePairEdgeList[i].pair2
                    && AttributePairEdgeList[i].edgepotentialfeaturevector.HasEdge == true
                    && AttributePairEdgeList[i].pair1.indexParent == OldParentRowNum)
                {
                    this.RepairSinglePair(AttributePairList, OldParentRowNum, AttributePairEdgeList[i].pair1.indexChild, NewParentRowNum);
                }
            }
        }

        /// <summary>
        /// Used only for the inner construction
        /// </summary>
        /// <param name="AttributePairList"></param>
        /// <param name="OldParentRowNum"></param>
        /// <param name="TargetRowNum"></param>
        /// <param name="NewParentRowNum"></param>
        /// <returns></returns>
        private int RepairSinglePair(IList<AnotationPair> AttributePairList, int OldParentRowNum, int TargetRowNum, int NewParentRowNum)
        {
            int index = 0;
            for (int i = 0; i < AttributePairList.Count; i++)
            {

                if (AttributePairList[i].indexChild == TargetRowNum && AttributePairList[i].indexParent == OldParentRowNum)
                {
                    index = i;
                    AttributePairList[i].nodepotentialfeaturevector.label = false;
                    break;
                }
                else if (AttributePairList[i].indexChild == TargetRowNum && AttributePairList[i].indexParent == NewParentRowNum)
                {
                    AttributePairList[i].nodepotentialfeaturevector.label = true;
                }
            }
            return index;
        }

    }
}
