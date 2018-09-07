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
            new Excelcol("Primer Apellido", typeof(string)),
            new Excelcol("Segundo Apellido", typeof(string)),
            new Excelcol("Nombres", typeof(string)),
            new Excelcol("Apellido Casada", typeof(string)),
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
        private Dist_File file;
        public PayrollExcel(Stream data, ApplicationDbContext context, string fileName, string mes, string gestion, string segmentoOrigen,Dist_File file, int headerin = 1, int sheets = 1, string resultfileName = "PayrollResult")
            : base(cols, data, fileName, headerin: headerin, resultfileName: resultfileName,sheets:sheets)
        {
            this.segmentoOrigen = segmentoOrigen;
            this.gestion = gestion;
            this.mes = mes;
            this.file = file;
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

            if (!connB1.connectedtoHana)
            {
                addError("Error en SAP", "No se puedo conectar con SAP B1, es posible que algunas validaciones cruzadas con SAP no sean ejecutadas");
            }

            //bool v1 = VerifyColumnValueIn(13, connB1.getBusinessPartners().ToList(), comment: "Esta AFP no esta registrada como un Bussines Partner en SAP");
            bool v2 = VerifyColumnValueIn(20, _context.TipoEmpleadoDists.Select(x => x.Name).ToList(), comment: "Este Tipo empleado no es valido.\n");
            var xxx = connB1.getCostCenter(B1Connection.Dimension.PEI, mes: this.mes, gestion: this.gestion)
                .Cast<string>().ToList();
            bool v3 = VerifyColumnValueIn(21, connB1.getCostCenter(B1Connection.Dimension.PEI,mes:this.mes,gestion:this.gestion).Cast<string>().ToList(), comment: "Este PEI no existe en SAP.\n");
            bool v4 = VerifyColumnValueIn(23, _context.Dependencies.Select(m => m.Cod).Distinct().ToList(),comment:"Esta Dependencia no existe en la Base de Datos Nacional.\n");
            bool v5 = VerifyPerson(ci: 1, CUNI: 19, fullname: 2, personActive: true, branchesId:Int32.Parse(this.segmentoOrigen), date: this.gestion + "-" + this.mes + "-01", dependency:23,paintdep:true,tipo:20);
            //bool v6 = VerifyColumnValueIn(25, connB1.getBusinessPartners(),comment:"Este seguro no esta registrado como un Bussines Partner en SAP");
            bool v7 = ValidateLiquidoPagable();
            bool v8 = ValidatenoZero();
            return isValid() /*&& v1*/ && v2 && v3 && v4 && v5 /*&& v6*/  && v7 && v8;
        }

        public bool ValidatenoZero(int sheet = 1)
        {
            int tipo = 20;
            int nozero = 22;
            bool res = true;
            string comment = "Este Valor no puede se cero.";
            IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();
            var l = UsedRange.LastRow().RowNumber();
            for (int i = headerin + 1; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                int nz = -1;
                if (wb.Worksheet(sheet).Cell(i, tipo).Value.ToString() != "TH" && Int32.TryParse(wb.Worksheet(sheet).Cell(i, nozero).Value.ToString(),out nz))
                {
                    if (nz == 0)
                    {
                        res = false;
                        paintXY(nozero, i, XLColor.Red, comment);
                    }
                }
            }
            valid = valid && res;
            return res;
        }

        public bool ValidateLiquidoPagable(int sheet=1)
        {
            int i1 = 6, i2 = 7, i3 = 8, i4 = 9, i5 = 10, i6 =11;
            int d1 = 14, d2 = 15, d3 = 16;
            int errortg = 12, errortd = 17, errorlp = 18;

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

                var ingresos = Math.Round(in1, 2) + Math.Round(in2, 2) + Math.Round(in3, 2) + Math.Round(in4, 2) + Math.Round(in5, 2) + Math.Round(in6, 2);
                var descuentos = Math.Round(de1, 2) + Math.Round(de2, 2) + Math.Round(de3, 2);

                Decimal.TryParse(wb.Worksheet(sheet).Cell(i, 12).Value.ToString(), out ti);
                Decimal.TryParse(wb.Worksheet(sheet).Cell(i, 17).Value.ToString(), out td);
                Decimal.TryParse(wb.Worksheet(sheet).Cell(i, 18).Value.ToString(), out lp);


                if (ingresos != ti)
                {
                    res = false;
                    paintXY(errortg, i, XLColor.Red, "No cuadran Ingresos. la suma sale: "+ingresos);
                    addError("No cuadran Ingresos", "La suma se calculó: " +ingresos + " se encontró " + ti);
                }

                if (descuentos != td)
                {
                    res = false;
                    paintXY(errortd, i, XLColor.Red, "No cuadran Descuentos. la suma sale: " + descuentos);
                    addError("No cuadran Descuentos", "La suma se calculó: " + descuentos + " se encontró " + td);
                }

                var dif = ingresos - descuentos;

                if ( dif!= lp)
                {
                    res = false;
                    paintXY(errorlp, i, XLColor.Red, "No cuadran Liquido Pagable. la suma sale: " + (ingresos-descuentos));
                    addError("No cuadran Liquido Pagable", "La suma se calculó: " + dif + " se encontró " + lp);
                }

            }

            valid = valid && res;
            return res;
        }

        public Dist_Payroll ToDistPayroll(int row, int sheet = 1)
        {
            Dist_Payroll payroll = new Dist_Payroll();
            payroll.Id = _context.Database.SqlQuery<int>("SELECT ADMNALRRHH.\"rrhh_Dist_Payroll_sqs\".nextval FROM DUMMY;").ToList()[0];
            payroll.Document = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
            payroll.FirstName = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            payroll.FirstSurName = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
            payroll.SecondSurName = wb.Worksheet(sheet).Cell(row, 4).Value.ToString();
            payroll.MariedSurName = wb.Worksheet(sheet).Cell(row, 5).Value.ToString();
            payroll.BasicSalary = strToDecimal(row, 6);
            payroll.AntiquityBonus = strToDecimal(row, 7);
            payroll.OtherIncome = strToDecimal(row, 8);
            payroll.TeachingIncome = strToDecimal(row, 9);
            payroll.OtherAcademicIncomes = strToDecimal(row, 10);
            payroll.Reintegro = strToDecimal(row, 11);
            payroll.TotalAmountEarned = strToDecimal(row, 12);
            payroll.AFP = wb.Worksheet(sheet).Cell(row, 13).Value.ToString();
            payroll.AFPLaboral = strToDecimal(row, 14);
            payroll.RcIva = strToDecimal(row, 15);
            payroll.Discounts = strToDecimal(row, 16);
            payroll.TotalAmountDiscounts = strToDecimal(row, 17);
            payroll.TotalAfterDiscounts = strToDecimal(row, 18);
            payroll.CUNI = wb.Worksheet(sheet).Cell(row, 19).Value.ToString();
            payroll.EmployeeType = wb.Worksheet(sheet).Cell(row, 20).Value.ToString();
            payroll.PEI = wb.Worksheet(sheet).Cell(row, 21).Value.ToString();
            payroll.WorkedHours = strToDouble(row, 22);
            payroll.Dependency = wb.Worksheet(sheet).Cell(row, 23).Value.ToString();
            payroll.AFPPatronal = strToDecimal(row, 24);
            payroll.IdentificadorSSU = wb.Worksheet(sheet).Cell(row, 25).Value.ToString();
            payroll.SeguridadCortoPlazoPatronal = strToDecimal(row, 26);
            payroll.ProvAguinaldo = strToDecimal(row, 27);
            payroll.ProvPrimas = strToDecimal(row, 28);
            payroll.ProvIndeminizacion = strToDecimal(row, 29);
            payroll.ProcedureTypeEmployee = payroll.EmployeeType;

            payroll.Porcentaje = 0m;
            payroll.mes = this.mes;
            payroll.gestion = this.gestion;
            payroll.segmentoOrigen = this.segmentoOrigen;
            payroll.DistFileId = file.Id;
            return payroll;
        }
    }
}