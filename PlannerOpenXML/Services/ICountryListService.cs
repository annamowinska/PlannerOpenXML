using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlannerOpenXML.Services
{
    public interface ICountryListService
    {
        public List<string> GetCountryCodes();

        public void UpdateCountryCodes(List<string> newCountryCodes);
    }
}
