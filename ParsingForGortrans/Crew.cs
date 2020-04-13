using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParsingForGortrans
{
    public class Crew
    {
        public int Number { get; set; }
        public List<Pair> Pairs { get;}

        public Crew(int number)
        {
            Number = number;
            Pairs = new List<Pair>();
        }

        public void SetListPair(List<Pair> pairs)
        {
            Pairs.AddRange(pairs);
        }
    }
}
