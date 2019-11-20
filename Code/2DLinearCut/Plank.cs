using System.Collections.Generic;
using System.Linq;

namespace _2DLinearCut
{
    class Plank
    {
        public double plankLength;

        public List<double> Cuts = new List<double>();

        public Plank(double length)
        {
            plankLength = length;
        }

        public double lengthRem()
        {
            return plankLength - Cuts.Sum();
        }
       

        public void Cut(double cutLength)
        {
            Cuts.Add(cutLength);
        }
    }
}
