using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOM_Version.Services
{
    public interface IDataService
    {
        bool Save(out string message);
        void Update(string codeHeader);
        bool LoadHeader(string code);
        bool LoadDetail(string code);
        string FindLastByItemCode(string itemCode);
    }
}
