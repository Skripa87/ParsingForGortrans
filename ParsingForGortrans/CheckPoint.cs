using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParsingForGortrans
{
    public class CheckPoint
    {
        public string Name { get; set; }
        public TimeSpan Time { get; set; }
        public bool IsEndpoint { get; set; }
        public int PitStopTimeStart { get; set; }
        public int PitStopTimeEnd { get; set; }
    }
}
