using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using UcbBack.Controllers;
using UcbBack.Logic.B1;
using UcbBack.Models;

namespace UcbBack.Logic.ExcelFiles
{
    public class PayrollExcel:ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("Carnet Identidad", typeof(string)), 
            new Excelcol("Nombre Completo", typeof(string)),
            new Excelcol("Haber Basico", typeof(double)),
            new Excelcol("Bono de Antigüedad", typeof(double)),
            new Excelcol("Otros Ingresos", typeof(double)),
            new Excelcol("Ingresos por docencia", typeof(double)),
            new Excelcol("Ingresos por otras actividades academicas", typeof(double)),
            new Excelcol("Total Ganado", typeof(double)),
            new Excelcol("Identificador de AFP", typeof(string)),
            new Excelcol("Aporte Laboral AFP", typeof(double)),
            new Excelcol("RC-IVA", typeof(double)),
            new Excelcol("Descuentos", typeof(double)),
            new Excelcol("Anticipos", typeof(double)),
            new Excelcol("Total Deducciones", typeof(double)),
            new Excelcol("Liquido Pagable", typeof(double)),
            new Excelcol("Segmento", typeof(string)),
            new Excelcol("Codigo nacional RRHH", typeof(string)),
            new Excelcol("Tipo empleado", typeof(string)),
            new Excelcol("PEI", typeof(string)),
            new Excelcol("Horas de trabajo mensual", typeof(double)),
            new Excelcol("Fecha Nacimiento", typeof(string)),
            new Excelcol("Unidad Organigrama", typeof(string)),
            new Excelcol("Aporte Patronal AFP", typeof(double)),
            new Excelcol("Aporte Patronal Seguridad Corto Plazo", typeof(double)),
            new Excelcol("Provision para Aguinaldos", typeof(double)),
            new Excelcol("Provision para Primas", typeof(double)),
            new Excelcol("Provisión para Indeminizacion", typeof(double)),
            new Excelcol("Unidad Organizacional SAP", typeof(string))
        };
        private ApplicationDbContext _context;
        public PayrollExcel(Stream data,ApplicationDbContext context, string fileName, int headerin =1, int sheets = 1, string resultfileName = "PayrollResult")
            : base(cols, data, fileName, headerin: headerin, resultfileName: resultfileName,sheets:sheets)
        {
            _context = context;
            isFormatValid();
        }

        public override void toDataBase()
        {
            IXLRange UsedRange = wb.Worksheet(1).RangeUsed();

            for (int i = 1 + headerin; i <= UsedRange.RowCount() + headerin; i++)
            {
                _context.DistPayrolls.Add(ToDistPayroll(i));
            }

            _context.SaveChanges();
        }

        public override bool ValidateFile()
        {
            var connB1 = B1Connection.Instance;

            bool v3 = VerifyColumnValueIn(22, _context.Dependencies.Select(m => m.Name).Distinct().ToList(),comment:"Esta Dependencia no existe en la Base de Datos Nacional.");
            bool v4 = VerifyColumnValueIn(28, _context.OrganizationalUnits.Select(m => m.Name).Distinct().ToList(),comment:"Esta Unidad Organigrama no existe en la Base de Datos Nacional.");
            bool v5 = VerifyPerson(1, 17, 2);
            bool v2 = VerifyColumnValueIn(28, connB1.getCostCenter(B1Connection.Dimension.OrganizationalUnit,col:"PrcName").ToList(), comment: "Esta Unidad Organigrama no existe en SAP.");
            return isValid() && v3 && v4 && v5 && v2;
        }

        public Dist_Payroll ToDistPayroll(int row, int sheet = 1)
        {
            Dist_Payroll payroll = new Dist_Payroll();
            payroll.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Dist_Academic_sqs\".nextval FROM DUMMY;").ToList()[0];
            payroll.Document = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
            payroll.FullName = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            payroll.BasicSalary = wb.Worksheet(sheet).Cell(row, 3).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 3).Value.ToString());
            payroll.AntiquityBonus = wb.Worksheet(sheet).Cell(row, 4).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 4).Value.ToString());
            payroll.OtherIncome = wb.Worksheet(sheet).Cell(row, 5).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 5).Value.ToString());
            payroll.TeachingIncome = wb.Worksheet(sheet).Cell(row, 6).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 6).Value.ToString());
            payroll.OtherAcademicIncomes = wb.Worksheet(sheet).Cell(row, 7).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 7).Value.ToString());
            payroll.TotalAmountEarned = wb.Worksheet(sheet).Cell(row, 8).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 8).Value.ToString());
            payroll.AFP = wb.Worksheet(sheet).Cell(row, 9).Value.ToString();
            payroll.AFPLaboral = wb.Worksheet(sheet).Cell(row, 10).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 10).Value.ToString());
            payroll.RcIva = wb.Worksheet(sheet).Cell(row, 11).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 11).Value.ToString());
            payroll.Discounts = wb.Worksheet(sheet).Cell(row, 12).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 12).Value.ToString());
            payroll.PaymentAdvances = wb.Worksheet(sheet).Cell(row, 13).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 13).Value.ToString());
            payroll.TotalAmountDiscounts = wb.Worksheet(sheet).Cell(row, 14).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 14).Value.ToString());
            payroll.TotalAfterDiscounts = wb.Worksheet(sheet).Cell(row, 15).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 15).Value.ToString());
            payroll.Segment = wb.Worksheet(sheet).Cell(row, 16).Value.ToString();
            payroll.CUNI = wb.Worksheet(sheet).Cell(row, 17).Value.ToString();
            payroll.EmployeeType = wb.Worksheet(sheet).Cell(row, 18).Value.ToString();
            payroll.PEI = wb.Worksheet(sheet).Cell(row, 19).Value.ToString();
            payroll.WorkedHours = wb.Worksheet(sheet).Cell(row, 20).Value.ToString() == "" ? 0.0 : Double.Parse(wb.Worksheet(sheet).Cell(row, 20).Value.ToString());
            payroll.BirthDate = wb.Worksheet(sheet).Cell(row, 21).Value.ToString();
            payroll.Dependency = wb.Worksheet(sheet).Cell(row, 22).Value.ToString();
            payroll.AFPPatronal = wb.Worksheet(sheet).Cell(row, 23).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 23).Value.ToString());
            payroll.SeguridadCortoPlazoPatronal = wb.Worksheet(sheet).Cell(row, 24).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 24).Value.ToString());
            payroll.ProvAguinaldo = wb.Worksheet(sheet).Cell(row, 25).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 25).Value.ToString());
            payroll.ProvPrimas = wb.Worksheet(sheet).Cell(row, 26).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 26).Value.ToString());
            payroll.ProvIndeminizacion = wb.Worksheet(sheet).Cell(row, 27).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 27).Value.ToString());
            payroll.SAPOrganizationalUnit = wb.Worksheet(sheet).Cell(row, 28).Value.ToString();
            payroll.Matched = 0;
            payroll.ProcedureTypeEmployee = "";

            return payroll;
        }
    }
}