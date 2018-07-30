using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Web;

namespace UcbBack.Models
{
    [NotMapped]
    public class Distribution
    {
        public string Document { get; set; }
        public string TipoEmpleado { get; set; }
        public string Dependency { get; set; }
        public string PEI { get; set; }
        public string PlanEstudios { get; set; }
        public string Paralelo { get; set; }
        public string Periodo { get; set; }
        public string Project { get; set; }
        public string Monto { get; set; }
        public string Porcentaje { get; set; }
        public string MontoDividido { get; set; }
        public string segmentoOrigen { get; set; }
        public string mes { get; set; }
        public string gestion { get; set; }
        public string Branches { get; set; }
        public string Concept { get; set; }
        public string CuentasContables { get; set; }
        public string Indicator { get; set; }


        public DataTable CreateDataTable<T>(IEnumerable<T> list)
        {
            Type type = typeof(T);
            var properties = type.GetProperties();

            DataTable dataTable = new DataTable();
            foreach (PropertyInfo info in properties)
            {
                dataTable.Columns.Add(new DataColumn(info.Name, Nullable.GetUnderlyingType(info.PropertyType) ?? info.PropertyType));
            }

            foreach (T entity in list)
            {
                object[] values = new object[properties.Length];
                for (int i = 0; i < properties.Length; i++)
                {
                    values[i] = properties[i].GetValue(entity);
                }

                dataTable.Rows.Add(values);
            }

            return dataTable;
        }
    }
}