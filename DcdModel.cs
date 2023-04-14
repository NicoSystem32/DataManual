using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Programa
{
    public class DCDModel
    {
        public int DCD_CODIGO { get; set; }
        public string DCD_CONSECUTIVO { get; set; }
        public DateTime? DCD_FECHA_EXPEDICION { get; set; }
        public int? DCD_CONVENIO_CODIGO { get; set; }
        public int? DCD_CONVENIO_TIPO { get; set; }
        public decimal? DCD_VALOR_ASEGURADO { get; set; }
        public int? DCD_ENTIDAD_CODIGO_PN { get; set; }
        public string DCD_DECFEP_CODIGO { get; set; }
        public int? DCD_PRODUCTOPARA { get; set; }
        public string DCD_VALIDO { get; set; }
        public string DCD_NOMBRE_REPLEGAL { get; set; }
        public string DCD_CEDULA_REPLEGAL { get; set; }
        public string DCD_BLOQUEOEDICION { get; set; }
        public string DCD_PERIODO_ANO { get; set; }
        public string DCD_PERIODO_MES { get; set; }
        public int? DCD_ID_MERCADO { get; set; }
        public int? DCD_ID_ESTADO { get; set; }
        public int? DCD_PROVEEDOR_CODIGO { get; set; }
        public decimal? DCD_KG { get; set; }
        public int? DCD_TIPO_PRODUCTO { get; set; }
        public decimal? DCD_KG_DEMOSTRADOS { get; set; }
        public int? DCD_TIPO_DOCUMENTO { get; set; }
        public int? DCD_CLASE_DOCUMENTO { get; set; }
        public int? DCD_PROCESO_INSERCION { get; set; }
        public int? DCD_KG_DEMOSTRADOS_POLIZA { get; set; }
        public string POLIZANUMERO { get; set; }
    }

}
