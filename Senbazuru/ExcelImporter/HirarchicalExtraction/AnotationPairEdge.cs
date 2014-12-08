using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Senbazuru.HirarchicalExtraction
{
    public class AnotationPairEdge
    {
        public AnotationPair pair1;
        public AnotationPair pair2;

        public EdgePotentialFeatureVector edgepotentialfeaturevector = null;

        public AnotationPairEdge(AnotationPair pair1, AnotationPair pair2)
        {
            this.pair1 = pair1;
            this.pair2 = pair2;
        }
    }
}
