using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Senbazuru.HirarchicalExtraction
{
    public class AnotationPair
    {
        public IList<Range> CellList = null;
        public int indexParent = 0;
        public int indexChild = 0;

        // Feature Vector denotes the list of feature values.
        public NodePotentialFeatureVector nodepotentialfeaturevector = null ;

        public AnotationPair(IList<Range> CellList, int indexParent, int indexChild)
        {
            this.CellList = CellList;
            this.indexParent = indexParent;
            this.indexChild = indexChild;
        }

        public static bool operator == (AnotationPair pair1, AnotationPair pair2)
        {
            return pair1.indexChild == pair2.indexChild && pair1.indexParent == pair2.indexParent ? true : false;
        }

        public static bool operator !=(AnotationPair pair1, AnotationPair pair2)
        {
            return pair1.indexChild != pair2.indexChild || pair1.indexParent == pair2.indexParent ? true : false;
        }
    }
}
