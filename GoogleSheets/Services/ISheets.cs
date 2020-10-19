using System.Collections.Generic;
using System.Threading.Tasks;

namespace GoogleSheets.Services
{
    public interface ISheets
    {
       List<string> ReturnName(string number);
    }
}
