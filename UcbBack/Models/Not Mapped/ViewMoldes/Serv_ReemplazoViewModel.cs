using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace UcbBack.Models.Not_Mapped.ViewMoldes
{
    public class Serv_ReemplazoViewModel
    {
        [DisplayName("Codigo Socio de Negocio")]
        public string Codigo_Socio_de_Negocio { get; set; }

        [DisplayName("Nombre Socio de Negocio")]
        public string Nombre_Socio_de_Negocio { get; set; }

        [DisplayName("Cod. Dependencia")]
        public string Cod_Dependencia { get; set; }

        [DisplayName("PEI-PO")]
        public string PEI_PO { get; set; }

        public string Glosa { get; set; }

        [DisplayName("Periodo Academico")]
        public string Periodo_Academico { get; set; }

        [DisplayName("Sigla Asignatura")]
        public string Sigla_Asignatura { get; set; }

        [DisplayName("Paralelo")]
        public string Paralelo { get; set; }

        [DisplayName("Código de Paralelo SAP")]
        public string Código_de_Paralelo_SAP { get; set; }

        [DisplayName("Tipo de Servicio")]
        public string Tipo_de_Servicio { get; set; }

        [DisplayName("Importe del Contrato")]
        public Decimal Importe_del_Contrato { get; set; }

        [DisplayName("Importe Deducción IUE")]
        public Decimal Importe_Deducción_IUE { get; set; }

        [DisplayName("Importe Deducción I.T.")]
        public Decimal Importe_Deducción_IT { get; set; }

        [DisplayName("Monto a Pagar")]
        public Decimal Monto_a_Pagar { get; set; }

        [DisplayName("Observaciones")]
        public string Observaciones { get; set; }
    }
}