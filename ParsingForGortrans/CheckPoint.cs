using DocumentFormat.OpenXml.Office2010.ExcelAc;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ParsingForGortrans
{
    public class CheckPoint: IComparable, IEquatable<CheckPoint>
    {
        public string Name { get; set; }
        public TimeSpan Time { get; set; }
        public bool IsEndpoint { get; set; }
        public TimeSpan PitStopTimeStart { get; set; }
        
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
                }
                else
                {
                    Time = TimeSpan.TryParse(bufferArray[0], out var result)
                         ? result
                         : TimeSpan.Zero;
                    PitStopTimeStart = TimeSpan.Zero;                    
                }
            }
            else 
            {
                Time = TimeSpan.TryParse(data, out var result)
                     ? result
                     : TimeSpan.Zero;
                PitStopTimeStart = TimeSpan.Zero;                
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

        public bool Equals(CheckPoint other)
        {
            if (other == null) return false;
            return string.Equals(Name.Trim(' ')
                                     .ToUpperInvariant(),
                           other.Name.Trim(' ')
                                     .ToUpperInvariant(), new StringComparison()) 
                && TimeSpan.Equals(PitStopTimeStart, other.PitStopTimeStart) &&
                               TimeSpan.Equals(Time, other.Time);                   
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

        public override int GetHashCode()
        {
            return Name.GetHashCode() & Time.GetHashCode() & PitStopTimeStart.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(this, obj))
            {
                return true;
            }

            if (ReferenceEquals(obj, null))
            {
                return false;
            }
            if (obj == null) return false;
            return string.Equals(((CheckPoint)obj).Name
                                    .Trim(' ')
                                    .ToUpperInvariant(), Name.Trim(' ').ToUpperInvariant(), new StringComparison())
                  && TimeSpan.Equals(PitStopTimeStart, ((CheckPoint)obj).PitStopTimeStart)
                  && TimeSpan.Equals(Time, ((CheckPoint)obj).Time);  
        }

        private CheckPoint() { }

        public static CheckPoint CreateEndPointFromCheckPointGroup(List<CheckPoint> checkPoints) 
        {
            if (checkPoints == null) return null;
            if (!checkPoints.Any()) return null;
            for (var i = 0;i < checkPoints.Count-1; i++)
            {
                if (!string.Equals(checkPoints[i].Name
                                                 .Trim(' ')
                                                 .ToUpperInvariant(), 
                                   checkPoints[i + 1].Name
                                                     .Trim(' ')
                                                     .ToUpperInvariant(), 
                                   new StringComparison())) return null;           
            }
            var max = checkPoints.Select(s => s.Time)
                                 .Max();
            var maxPitStopStart = checkPoints.Select(s => s.PitStopTimeStart)
                                             .Max();
            var buffArr = checkPoints.FindAll(f => f.PitStopTimeStart != TimeSpan.Zero);
            var minPitStopStart = buffArr.Any()
                                ? buffArr.Select(s => s.PitStopTimeStart)
                                 .Min()
                                : TimeSpan.MaxValue;
            buffArr = checkPoints.FindAll(f => f.Time != TimeSpan.Zero);
            var min = buffArr.Select(s => s.Time)
                                 .Min();
            return new CheckPoint()
            {
                IsEndpoint = true,
                Name = checkPoints.FirstOrDefault()
                                  .Name,
                Time = max > maxPitStopStart ? max : maxPitStopStart,
                PitStopTimeStart = min < minPitStopStart ? min : minPitStopStart
            };
        }
    }
}
