using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pdf_editor.classes
{
    public class Retorno<T>
    {
        public bool Success { get; set; }
        public int CodigoResposta { get; set; }
        public string Message { get; set; } = "";
        public List<T> Dados { get; set; } = new List<T>();

    }
}
