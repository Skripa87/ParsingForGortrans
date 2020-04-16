using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParsingForGortrans
{
    public class Pair
    {
        public int Number { get; set; }
        public TimeSpan StartWorkTime { get; set; }
        public TimeSpan EndWorkTime { get; set; }
        public TimeSpan DinnerStartTime { get; set; }
        public TimeSpan DinnerEndTime { get; set; }
        public TimeSpan StartSettling { get; set; }
        public TimeSpan EndSettling { get; set; }
        public List<Flight> Flights { get; }
        public int FlightsCount { get; set; }

        private void CreateWorkTime(List<string> data) 
        {
            string timeBuffer = "";
            try
            {
                timeBuffer = data.ElementAt(1);
            }
            catch (ArgumentOutOfRangeException)
            {
                timeBuffer = "";
            }
            var timeBufferArray = timeBuffer.Contains('-')
                                ? timeBuffer.Split('-')
                                : timeBuffer.Split(' ');
            if (timeBufferArray.Length == 0)
            {
                StartWorkTime = TimeSpan.MinValue;
                EndWorkTime = TimeSpan.MaxValue;
            }
            else if (timeBufferArray.Length == 1)
            {
                StartWorkTime = TimeSpan.TryParse(timeBufferArray[0], out var result)
                              ? result
                              : TimeSpan.MinValue;
                EndWorkTime = TimeSpan.MaxValue;
            }
            else
            {
                StartWorkTime = TimeSpan.TryParse(timeBufferArray[0], out var startresult)
                              ? startresult
                              : TimeSpan.MinValue;
                EndWorkTime = TimeSpan.TryParse(timeBufferArray[1], out var endresult)
                              ? endresult
                              : TimeSpan.MinValue;
            }
        }

        private void CreateDinnerOrSettlingTime(List<string> data) 
        {
            string timeBuffer = "";
            try
            {
                timeBuffer = data.ElementAt(2);
            }
            catch (ArgumentOutOfRangeException)
            {
                timeBuffer = "";
            }
            if (timeBuffer.Contains('-')) 
            {
                var timeBufferArray_ = timeBuffer.Trim()
                                                 .Split('-');
                switch (timeBufferArray_.Length) 
                {
                    case 0: StartSettling = TimeSpan.MinValue;
                            EndSettling = TimeSpan.MaxValue;  
                            break;
                    case 1: StartSettling = TimeSpan.TryParse(timeBufferArray_[0],out var result) 
                                          ? result
                                          : TimeSpan.MinValue;
                            EndSettling = TimeSpan.MaxValue;
                            break;
                    case 2: StartSettling = TimeSpan.TryParse(timeBufferArray_[0], out var startresult)
                                          ? startresult
                                          : TimeSpan.MinValue;
                            EndSettling = TimeSpan.TryParse(timeBufferArray_[1], out var endresult)
                                          ? endresult
                                          : TimeSpan.MaxValue; break;
                }
            }
            else 
            {
                var timeBufferArray = timeBuffer.Trim()
                                                .Split(' ');
                switch (timeBufferArray.Length)
                {
                    case 0:
                        DinnerStartTime = TimeSpan.MinValue;
                        DinnerEndTime = TimeSpan.MaxValue;
                        break;
                    case 1:
                        DinnerStartTime = TimeSpan.TryParse(timeBufferArray[0], out var result)
                                        ? result
                                        : TimeSpan.MinValue;
                        DinnerEndTime = TimeSpan.MaxValue;
                        break;
                    case 2:
                        DinnerStartTime = TimeSpan.TryParse(timeBufferArray[0], out var startresult)
                                        ? startresult
                                        : TimeSpan.MinValue;
                        DinnerEndTime = TimeSpan.TryParse(timeBufferArray[1], out var endresult)
                                      ? endresult
                                      : TimeSpan.MaxValue; break;
                }
            }            
        }        

        public Pair(List<string> data) 
        {
            if (data == null || data.Count == 0 || data.All(d => string.IsNullOrEmpty(d))) return;
            data.RemoveAll(d => string.IsNullOrEmpty(d));
            Number = int.TryParse(data?.FirstOrDefault() ?? "", out int number)
                   ? number
                   : 999;
            CreateWorkTime(data);
            CreateDinnerOrSettlingTime(data);
            FlightsCount = int.TryParse(data?.Last()?.Split(' ')[0], out var number1)
                         ? number1
                         : -999;
            Flights = new List<Flight>();
        }

        public void SetFligths(List<Flight> flights) 
        {
            Flights.AddRange(flights);
        }
    }


}
