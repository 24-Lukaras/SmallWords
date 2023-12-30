using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmallWords.Replacing
{
    public interface IDataReplacer
    {

        public void Replace(Stream stream, Dictionary<string, object> data);
    }
}
