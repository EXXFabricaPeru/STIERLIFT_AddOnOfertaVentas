using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddOnUI.Modelos
{
    public class OfertaVentaDatos
    {
        public int DocEntry { get; set; }
        public int DocNum { get; set; }
        public string CardCode { get; set; }
        public DateTime DocDate { get; set; }
        public DateTime DocDueDate { get; set; }
        public string OT { get; set; }
        public string Taller { get; set; }
        public string Sucursal { get; set; }    

        public List<DetalleOV> Lineas { get; set; } = new List<DetalleOV>();


    }

    public class DetalleOV
    {
        public string ItemCode { get; set; }
        public string Descripcion { get; set; }
        public string UoM { get; set; }
        public double Cantidad { get; set; }
        public string TaxCode { get; set; }
        public double Precio { get; set; }
        public string Almacen { get; set; }
        public string Proyecto { get; set; }
        public string Margen { get; set; }
        public string CentroCosto { get; set; }
        public string CentroBeneficio { get; set; }
        public string Equipos { get; set; }
        public string CC5 {get; set;}
        public string GrupoDetraccion { get; set; }
        public string GrupoRetencion { get; set; }
        public string NroOferta { get; set; }
        public string LineaOferta { get; set; }

    }
}
