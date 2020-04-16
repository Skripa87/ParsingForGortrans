using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParsingForGortrans
{
    public class CheckPoint: IComparable
    {
        public string Name { get; set; }
        public TimeSpan Time { get; set; }
        public bool IsEndpoint { get; set; }
        public TimeSpan PitStopTimeStart { get; set; }
        public TimeSpan PitStopTimeEnd { get; set; }

        public CheckPoint(string name, string data) 
        {
            Name = name;
            IsEndpoint = data.Contains('\n');
            if (IsEndpoint)
            {
                var bufferArray = data?.Split('\n') ?? new string[0];
                if(bufferArray.Length > 1)
                {
                    Time = TimeSpan.TryParse(bufferArray[1], out var result)
                         ? result
                         : TimeSpan.Zero;
                    PitStopTimeStart = TimeSpan.TryParse(bufferArray[0], out var startresult)
                                     ? startresult
                                     : TimeSpan.Zero;
                    PitStopTimeEnd = TimeSpan.TryParse(bufferArray[1], out var endresult)
                                     ? endresult
                                     : TimeSpan.Zero;
                }
                else
                {
                    Time = TimeSpan.TryParse(bufferArray[0], out var result)
                         ? result
                         : TimeSpan.Zero;
                    PitStopTimeStart = TimeSpan.Zero;
                    PitStopTimeEnd = TimeSpan.Zero;
                }
            }
            else 
            {
                Time = TimeSpan.TryParse(data, out var result)
                     ? result
                     : TimeSpan.Zero;
                PitStopTimeStart = TimeSpan.Zero;
                PitStopTimeEnd = TimeSpan.Zero;
            }
        }

        public int CompareTo(object obj)
        {
            return ((CheckPoint)obj).Time > Time
                ? -1
                : (((CheckPoint)obj).Time < Time
                  ? 1 : 0);
        }
    }
}
