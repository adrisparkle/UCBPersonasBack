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
            new Excelcol("Haber Básico", typeof(double)),
            new Excelcol("Bono de Antigüedad", typeof(double)),
            new Excelcol("Otros Ingresos", typeof(double)),
            new Excelcol("Ingresos por docencia", typeof(double)),
            new Excelcol("Ingresos por otras actividades académicas", typeof(double)),
            new Excelcol("Reintegro", typeof(double)),
            new Excelcol("Total Ganado", typeof(double)),
            new Excelcol("Identificador de AFP", typeof(string)),
            new Excelcol("Aporte Laboral AFP", typeof(double)),
            new Excelcol("RC-IVA", typeof(double)),
            new Excelcol("Descuentos", typeof(double)),
            new Excelcol("Total Deducciones", typeof(double)),
            new Excelcol("Liquido Pagable", typeof(double)),
            new Excelcol("CUNI", typeof(string)),
            new Excelcol("Tipo empleado", typeof(string)),
            new Excelcol("PEI", typeof(string)),
            new Excelcol("Horas de trabajo mensual", typeof(double)),
            new Excelcol("Identificador de Dependencia", typeof(string)),
            new Excelcol("Aporte Patronal AFP", typeof(double)),
            new Excelcol("Identificador Seguridad Corto Plazo", typeof(string)),
            new Excelcol("Aporte Patronal SCP", typeof(double)),
            new Excelcol("Provisión Aguinaldos", typeof(double)),
            new Excelcol("Provisión Primas", typeof(double)),
            new Excelcol("Provisión Indemnización", typeof(double))
        };

        private string mes, gestion, segmentoOrigen;
        private ApplicationDbContext _context;
        public PayrollExcel(Stream data, ApplicationDbContext context, string fileName, string mes, string gestion, string segmentoOrigen, int headerin = 1, int sheets = 1, string resultfileName = "PayrollResult")
            : base(cols, data, fileName, headerin: headerin, resultfileName: resultfileName,sheets:sheets)
        {
            this.segmentoOrigen = segmentoOrigen;
            this.gestion = gestion;
            this.mes = mes;
            _context = context;
            isFormatValid();
        }
        public PayrollExcel(string fileName,int headerin = 1):base(cols,fileName,headerin)
        { }

        public override void toDataBase()
        {
            IXLRange UsedRange = wb.Worksheet(1).RangeUsed();

            for (int i = 1 + headerin; i <= UsedRange.LastRow().RowNumber() ; i++)
            {
                _context.DistPayrolls.Add(ToDistPayroll(i));
            }

            _context.SaveChanges();
        }

        public override bool ValidateFile()
        {
            var connB1 = B1Connection.Instance;
            //todo cabiar a socio de negocio
            bool v1 = VerifyColumnValueIn(10,new List<string>{"FUT","PRE"});
            bool v2 = VerifyColumnValueIn(17, _context.TipoEmpleadoDists.Select(x => x.Name).ToList(), comment: "Este Tipo empleado no es valido.");
            bool v3 = VerifyColumnValueIn(18, connB1.getCostCenter(B1Connection.Dimension.PEI,mes:this.mes,gestion:this.gestion).Cast<string>().ToList(), comment: "Este PEI no existe en SAP.");
            bool v4 = VerifyColumnValueIn(20, _context.Dependencies.Select(m => m.Cod).Distinct().ToList(),comment:"Esta Dependencia no existe en la Base de Datos Nacional.");
            bool v5 = VerifyPerson(ci:1, CUNI:16, fullname:2);
            bool v6 = VerifyColumnValueIn(22, connB1.getBusinessPartners(),comment:"Este seguro no esta registrado como un Bussines Partner en SAP");
            bool v7 = ValidateLiquidoPagable();
            return isValid() && v1 && v2 && v3 && v4 && v5 && v6 && v7;
        }

        public bool ValidateLiquidoPagable(int sheet=1)
        {
            int i1 = 3, i2 = 4, i3 = 5, i4 = 6, i5 = 7, i6 =8;
            int d1 = 11, d2 = 12, d3 = 13;

            bool res = true;
            IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();

            for (int i = headerin + 1; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                decimal in1 = 0, in2 = 0, in3 = 0, in4 = 0, in5 = 0, in6 = 0,ti=0,lp=0;
                decimal de1 = 0, de2 = 0, de3 = 0, td=0;
                Decimal.TryParse(wb.Worksheet(sheet).Cell(i, i1).Value.ToString(), out in1);
                Decimal.TryParse(wb.Worksheet(sheet).Cell(i, i2).Value.ToString(), out in2);
                Decimal.TryParse(wb.Worksheet(sheet).Cell(i, i3).Value.ToString(), out in3);
                Decimal.TryParse(wb.Worksheet(sheet).Cell(i, i4).Value.ToString(), out in4);
                Decimal.TryParse(wb.Worksheet(sheet).Cell(i, i5).Value.ToString(), out in5);
                Decimal.TryParse(wb.Worksheet(sheet).Cell(i, i6).Value.ToString(), out in6);

                Decimal.TryParse(wb.Worksheet(sheet).Cell(i, d1).Value.ToString(), out de1);
                Decimal.TryParse(wb.Worksheet(sheet).Cell(i, d2).Value.ToString(), out de2);
                Decimal.TryParse(wb.Worksheet(sheet).Cell(i, d3).Value.ToString(), out de3);

                var ingresos = in1 + in2 + in3 + in4 + in5 + in6;
                var descuentos = de1 + de2 + de3;

                Decimal.TryParse(wb.Worksheet(sheet).Cell(i, 9).Value.ToString(), out ti);
                Decimal.TryParse(wb.Worksheet(sheet).Cell(i, 14).Value.ToString(), out td);
                Decimal.TryParse(wb.Worksheet(sheet).Cell(i, 15).Value.ToString(), out lp);


                if (ingresos != ti)
                {
                    res = false;
                    paintXY(9, i, XLColor.Red, "No cuadran Ingresos. la suma sale: "+ingresos);
                }

                if (descuentos != td)
                {
                    res = false;
                    paintXY(14, i, XLColor.Red, "No cuadran Descuentos. la suma sale: " + descuentos);
                }

                var dif = ingresos - descuentos;

                if ( dif!= lp)
                {
                    res = false;
                    paintXY(15, i, XLColor.Red, "No cuadran Liquido Pagable. la suma sale: " + (ingresos-descuentos));
                }

            }

            return res;
        }

        public Dist_Payroll ToDistPayroll(int row, int sheet = 1)
        {
            Dist_Payroll payroll = new Dist_Payroll();
            payroll.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Dist_Payroll_sqs\".nextval FROM DUMMY;").ToList()[0];
            payroll.Document = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
            payroll.FullName = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            payroll.BasicSalary = strToDecimal(row, 3);
            payroll.AntiquityBonus = strToDecimal(row, 4);
            payroll.OtherIncome = strToDecimal(row, 5);
            payroll.TeachingIncome = strToDecimal(row, 6);
            payroll.OtherAcademicIncomes = strToDecimal(row, 7);
            payroll.Reintegro = strToDecimal(row, 8);
            payroll.TotalAmountEarned = strToDecimal(row, 9);
            payroll.AFP = wb.Worksheet(sheet).Cell(row, 10).Value.ToString();
            payroll.AFPLaboral = strToDecimal(row, 11);
            payroll.RcIva = strToDecimal(row, 12);
            payroll.Discounts = strToDecimal(row, 13);
            payroll.TotalAmountDiscounts = strToDecimal(row, 14);
            payroll.TotalAfterDiscounts = strToDecimal(row, 15);
            payroll.CUNI = wb.Worksheet(sheet).Cell(row, 16).Value.ToString();
            payroll.EmployeeType = wb.Worksheet(sheet).Cell(row, 17).Value.ToString();
            payroll.PEI = wb.Worksheet(sheet).Cell(row, 18).Value.ToString();
            payroll.WorkedHours = strToDouble(row, 19);
            payroll.Dependency = wb.Worksheet(sheet).Cell(row, 20).Value.ToString();
            payroll.AFPPatronal = strToDecimal(row, 21);
            payroll.IdentificadorSSU = wb.Worksheet(sheet).Cell(row, 22).Value.ToString();
            payroll.SeguridadCortoPlazoPatronal = strToDecimal(row, 23);
            payroll.ProvAguinaldo = strToDecimal(row, 24);
            payroll.ProvPrimas = strToDecimal(row, 25);
            payroll.ProvIndeminizacion = strToDecimal(row, 26);
            payroll.ProcedureTypeEmployee = "";

            payroll.Porcentaje = 0m;
            payroll.mes = this.mes;
            payroll.gestion = this.gestion;
            payroll.segmentoOrigen = this.segmentoOrigen;
            return payroll;
        }
    }
}