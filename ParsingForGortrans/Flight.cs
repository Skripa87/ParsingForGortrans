using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParsingForGortrans
{
    public class Flight
    {
        public List<CheckPoint> CheckPoints { get; }
        public int Number { get; set; }

        public Flight(int number)
        {
            Number = number;
            CheckPoints = new List<CheckPoint>();
        }

        public void InitCheckPoints(List<CheckPoint> checkPoints)
        {
            CheckPoints.AddRange(checkPoints);
        } 
    }
}
