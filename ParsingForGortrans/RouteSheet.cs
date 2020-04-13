using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParsingForGortrans
{
    public class RouteSheet
    {
        public string ShortName { get; set; }
        public string FullName { get; set; }
        public List<Crew> Crews { get;}

        public RouteSheet(string shortName, string fullName)
        {
            ShortName = shortName;
            FullName = fullName;
            Crews = new List<Crew>();
        }

        public void InitCrews(List<Crew> crews)
        {
            Crews.AddRange(crews);
        }
    }    
}
