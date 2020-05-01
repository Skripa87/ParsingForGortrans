using System;
using System.Linq;

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
                if(bufferArray.Length > 1 && !string.IsNullOrWhiteSpace(bufferArray[1]))
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
            return obj == null 
                  ? 1 
                  : (((CheckPoint)obj).Time > Time
                     ? -1
                     : (((CheckPoint)obj).Time < Time
                        ? 1 
                        : 0));
        }

        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            return ((CheckPoint)obj).IsEndpoint == IsEndpoint
                && string.Equals(((CheckPoint)obj).Name.ToUpperInvariant(), Name.ToUpperInvariant(), new StringComparison())
                && ((CheckPoint)obj).Time == Time
                && ((CheckPoint)obj).PitStopTimeEnd == PitStopTimeEnd
                && ((CheckPoint)obj).PitStopTimeStart == PitStopTimeStart;
        }        

        public static bool operator ==(CheckPoint left, CheckPoint right)
        {
            if (ReferenceEquals(left, null))
            {
                return ReferenceEquals(right, null);
            }
            return left.Equals(right);
        }

        public static bool operator !=(CheckPoint left, CheckPoint right)
        {
            return !(left == right);
        }

        public static bool operator <(CheckPoint left, CheckPoint right)
        {
            return ReferenceEquals(left, null) ? !ReferenceEquals(right, null) : left.CompareTo(right) < 0;
        }

        public static bool operator <=(CheckPoint left, CheckPoint right)
        {
            return ReferenceEquals(left, null) || left.CompareTo(right) <= 0;
        }

        public static bool operator >(CheckPoint left, CheckPoint right)
        {
            return !ReferenceEquals(left, null) && left.CompareTo(right) > 0;
        }

        public static bool operator >=(CheckPoint left, CheckPoint right)
        {
            return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.CompareTo(right) >= 0;
        }
    }
}
