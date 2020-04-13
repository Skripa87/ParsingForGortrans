using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParsingForGortrans
{
    public class Pair
    {
        public TimeSpan StartWorkTime { get; set; }
        public TimeSpan EndWorkTime { get; set; }
        public TimeSpan DinnerStartTime { get; set; }
        public TimeSpan DinnerEndTime { get; set; }
        public List<CircleRoute> CircleRoutes { get; set; }
        public int CircleRoutesCount { get; set; }
    }
}
