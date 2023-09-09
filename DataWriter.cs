using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileWriter
{
    internal interface DataWriter
    {
        void Write(List<Object> objects);
    }
}
