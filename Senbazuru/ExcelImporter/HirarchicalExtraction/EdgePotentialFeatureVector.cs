using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Senbazuru.HirarchicalExtraction
{
    public class EdgePotentialFeatureVector
    {
        public IList<int> features = new List<int>();

        public bool HasEdge = false;

        public int pair1Idx;
        public int pair2Idx;

        public EdgePotentialFeatureVector(IList<int> features, int pair1, int pair2)
        {
            this.features = features ;
            this.pair1Idx = pair1;
            this.pair2Idx = pair2;
        }
    }
}
