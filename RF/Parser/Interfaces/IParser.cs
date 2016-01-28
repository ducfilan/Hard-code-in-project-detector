using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RF.Parser
{
    public interface IParser
    {
        List<Item> GetHardCodeList(string filePath);
    }
}
