using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Programa
{
    public class FiniquitoManualModel
    {
        public int MAN_ID { get; set; }
        public int MAN_IDDCD { get; set; }
        public string? MAN_COMERCIALIZADORA { get; set; }
        public int MAN_ANIO { get; set; }
        public int MAN_MES { get; set; }
        public string? MAN_TIPO_DOCUMENTO { get; set; }
        public string? MAN_DOCUMENTO { get; set; }
        public decimal MAN_SUMA_KILPALMA { get; set; }
        public decimal MAN_SUMA_KILPALMISTE { get; set; }
        public string? MAN_PENDIENTE_FINIQUITO { get; set; }
        public int MAN_PROCESADO { get; set; }
        public DateTime? MAN_APROBADAFCP { get; set; }
    }
}
