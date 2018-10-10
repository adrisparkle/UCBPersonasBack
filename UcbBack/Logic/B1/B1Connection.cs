using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Web;
using Microsoft.Ajax.Utilities;
using Newtonsoft.Json.Linq;
using Sap.Data.Hana;
using SAPbobsCOM;
using UcbBack.Models;
using UcbBack.Models.Auth;
using UcbBack.Models.Dist;
using UcbBack.Models.Not_Mapped;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;
using Resource = SAPbobsCOM.Resource;

namespace UcbBack.Logic.B1
{
    public class B1Connection
    {
        private static object Lock = new Object();
        private static B1Connection instance=null;
        private struct BusinessObjectType
        {
            public static string BussinesPartner = "BUSINESSPARTNER";
            public static string Voucher = "VOUCHER";
            public static string Employee = "EMPLOYEE";
        }

        private SAPbobsCOM.Company company;
        private HanaConnection HanaConn;
        private int connectionResult;
        private int errorCode = 0;
        private string errorMessage = "";
        private string DatabaseName;
        public bool connectedtoHana=false;
        public bool connectedtoB1=false;
        private ApplicationDbContext _context;

        public enum Dimension
        {
            All,
            OrganizationalUnit,
            PEI,
            PlanAcademico,
            Paralelo,
            Periodo
        };


        private B1Connection()
        {
            try
            {
                DatabaseName = ConfigurationManager.AppSettings["HanaBD"];
                //string cadenadeconexion = "Server=192.168.18.180:30015;UserID=admnalrrhh;Password=Rrhh12345;Current Schema="+DatabaseName;
                //string cadenadeconexion = "Server=SAPHANA01:30015;UserID=SDKRRHH;Password=Rrhh1234;Current Schema=UCBTEST"+DatabaseName;
                string cadenadeconexion = "Server=" + ConfigurationManager.AppSettings["B1Server"] +
                                          ";UserID=" + ConfigurationManager.AppSettings["HanaBDUser"] +
                                          ";Password=" + ConfigurationManager.AppSettings["HanaPassword"] +
                                          ";Current Schema=" + ConfigurationManager.AppSettings["HanaBD"];
                HanaConn = new HanaConnection(cadenadeconexion);
                HanaConn.Open();
                connectedtoHana = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                connectedtoHana = false;
            }

            try
            {
                ConnectB1();
                connectedtoB1 = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                connectedtoB1 = false;
            }
            _context = new ApplicationDbContext();
            if (!connectedtoB1 && !connectedtoHana)
                instance = null;
        }

        // Double Check locking implementation for thread safe singleton
        public static B1Connection Instance()
        {
            if(instance == null) // 1st check
            {
                lock (Lock) // locked
                {
                    if (instance == null) // second check
                    {
                        instance = new B1Connection(); // instantiate a new (and the only one) instance
                    }
                }
            }

            return instance; // return the instance 
        }

        private bool DisconnectB1()
        {
            bool conectado = true;
            try
            {
                conectado = company.Connected;
                if (conectado)
                {
                    if (company.InTransaction)
                        company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                company.Disconnect();
                conectado = company.Connected;
            }
            catch
            { }
            return conectado;
        }

        private int ConnectB1()
        {
            company = new SAPbobsCOM.Company();

            /*company.Server = "SAPHANA01:30015";
            company.CompanyDB = "UCBTEST";
            company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
            company.DbUserName = "DESARROLLO1";
            company.DbPassword = "Rrhh12345";
            company.UserName = "manager7";
            company.Password = "sandra2018";
            company.language = SAPbobsCOM.BoSuppLangs.ln_English_Gb;
            company.UseTrusted = true;
            company.LicenseServer = "SAPHANA01:30015";
            company.SLDServer = "SAPHANA01:40000";*/

            company.Server = ConfigurationManager.AppSettings["B1Server"];
            company.CompanyDB = ConfigurationManager.AppSettings["B1CompanyDB"];
            company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
            company.DbUserName = ConfigurationManager.AppSettings["B1DbUserName"];
            company.DbPassword = ConfigurationManager.AppSettings["B1DbPassword"];
            company.UserName = ConfigurationManager.AppSettings["B1UserName"];
            company.Password = ConfigurationManager.AppSettings["B1Password"];
            company.language = SAPbobsCOM.BoSuppLangs.ln_English_Gb;
            company.UseTrusted = true;
            company.LicenseServer = ConfigurationManager.AppSettings["B1LicenseServer"];
            company.SLDServer = ConfigurationManager.AppSettings["B1SLDServer"];



            connectionResult = company.Connect();
            var x = company.Connected;
        if (connectionResult != 0)
            {
                company.GetLastError(out errorCode, out errorMessage);
            }
            return connectionResult;
        }  

        public bool TestHanaConection()
        {
            string cadenadeconexion = "Server=" + ConfigurationManager.AppSettings["B1Server"] +
                                      ";UserID=" + ConfigurationManager.AppSettings["HanaBDUser"] +
                                      ";Password=" + ConfigurationManager.AppSettings["HanaPassword"] +
                                      ";Current Schema=" + ConfigurationManager.AppSettings["HanaBD"];
            //Realizamos la conexion a SQL                            
            bool resultado = false;
            try
            {
                HanaConnection da = new HanaConnection(cadenadeconexion);
                da.Open();
                resultado = true;
                da.Close();
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                resultado = false;
            }
            return resultado;
        }

        private B1SDKLog initLog(int userId,string type,string ObjectId)
        {
            B1SDKLog log = new B1SDKLog();
            log.Id = B1SDKLog.GetNextId(_context);
            log.ObjectId = ObjectId;
            log.BusinessObject = type;
            log.Success = true;
            return log;
        }

        public string addPersonToEmployeeMasterData(int UserId,People person)
        {
            var log = initLog(UserId,BusinessObjectType.Employee,person.Id.ToString());
            try
            {
                if (company.Connected)
                {
                    company.StartTransaction();
                    SAPbobsCOM.EmployeesInfo oEmployeesInfo = (SAPbobsCOM.EmployeesInfo)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo);

                    oEmployeesInfo.FirstName = person.Names + "-4";
                    oEmployeesInfo.LastName = person.FirstSurName;
                    oEmployeesInfo.Gender = person.Gender == "M" ? BoGenderTypes.gt_Male : BoGenderTypes.gt_Female;
                    oEmployeesInfo.DateOfBirth = person.BirthDate;
                    oEmployeesInfo.ExternalEmployeeNumber = person.CUNI;

                    oEmployeesInfo.Add();
                    string newKey = company.GetNewObjectKey();
                    company.GetLastError(out errorCode, out errorMessage);
                    if (errorCode != 0)
                    {
                        log.Success = false;
                        log.ErrorCode = errorCode.ToString();
                        log.ErrorMessage = "SDK: " + errorMessage;
                        _context.SdkErrorLogs.Add(log);
                        _context.SaveChanges();
                        return "ERROR";
                    }
                    else
                    {
                        if (company.InTransaction)
                        {
                            company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            newKey = newKey.Replace("\t1", "");
                        }
                        _context.SdkErrorLogs.Add(log);
                        _context.SaveChanges();
                        return newKey;
                    }  
                }
                log.Success = false;
                log.ErrorMessage = "SDK: Not Connected";
                _context.SdkErrorLogs.Add(log);
                _context.SaveChanges();
                return "ERROR";
            }
            catch (Exception ex)
            {
                log.Success = false;
                log.ErrorMessage = "Catch: "+ex.Message;
                _context.SdkErrorLogs.Add(log);
                _context.SaveChanges();
                return "ERROR";
            }
        }

        public string updatePersonInBussinesPartner(int UserId, People person)
        {
            var log = initLog(UserId, BusinessObjectType.BussinesPartner, person.Id.ToString());
            try
            {
                if (company.Connected)
                {
                    company.StartTransaction();
                    SAPbobsCOM.BusinessPartners businessObject =
                        (SAPbobsCOM.BusinessPartners) company.GetBusinessObject(SAPbobsCOM.BoObjectTypes
                            .oBusinessPartners);
                    //if person exist as BusinesPartner
                    if (businessObject.GetByKey("R" + person.CUNI))
                    {
                        businessObject.CardName = person.FirstSurName + " " + person.Names;
                        businessObject.CardForeignName = person.FirstSurName + " " + person.Names;
                        businessObject.CardType = SAPbobsCOM.BoCardTypes.cCustomer;
                        businessObject.CardCode = "R" + person.CUNI;
                        businessObject.UserFields.Fields.Item("LicTradNum").Value = person.Document;
                        businessObject.GroupCode = 102;


                        // set Branch Code
                        businessObject.BPBranchAssignment.DisabledForBP = SAPbobsCOM.BoYesNoEnum.tNO;
                        businessObject.BPBranchAssignment.BPLID =
                            Int32.Parse(person.GetLastContract().Branches.CodigoSAP);
                        businessObject.BPBranchAssignment.Add();
                        // save new business partner
                        businessObject.Update();
                        // get the new code
                        string newKey = company.GetNewObjectKey();
                        company.GetLastError(out errorCode, out errorMessage);
                        if (errorCode != 0)
                        {
                            log.Success = false;
                            log.ErrorCode = errorCode.ToString();
                            log.ErrorMessage = "SDK: " + errorMessage;
                            _context.SdkErrorLogs.Add(log);
                            _context.SaveChanges();
                            return "ERROR";
                        }
                        else
                        {
                            if (company.InTransaction)
                            {
                                company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                newKey = newKey.Replace("\t1", "");
                                _context.SdkErrorLogs.Add(log);
                                _context.SaveChanges();
                                return newKey;
                            }
                        }
                    }
                    log.Success = false;
                    log.ErrorMessage = "SDK: Not Found";
                    _context.SdkErrorLogs.Add(log);
                    _context.SaveChanges();
                    return "ERROR";
                }
                log.Success = false;
                log.ErrorMessage = "SDK: Not Connected";
                _context.SdkErrorLogs.Add(log);
                _context.SaveChanges();
                return "ERROR";
            }
            catch (Exception ex)
            {
                log.Success = false;
                log.ErrorMessage = "Catch: " + ex.Message;
                _context.SdkErrorLogs.Add(log);
                _context.SaveChanges();
                return "ERROR";
            }
        }

        public string addpersonToBussinesPartner(int UserId, People person)
        {
            var log = initLog(UserId, BusinessObjectType.BussinesPartner, person.Id.ToString());
            try
            {
                if (company.Connected)
                {
                    company.StartTransaction();
                    SAPbobsCOM.BusinessPartners businessObject = (SAPbobsCOM.BusinessPartners)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);

                    businessObject.CardName = person.FirstSurName+ " " + person.Names;
                    businessObject.CardForeignName = person.FirstSurName+ " " + person.Names;
                    businessObject.CardType = SAPbobsCOM.BoCardTypes.cCustomer;
                    businessObject.CardCode = "R"+person.CUNI;

                    // set NIT
                    businessObject.UserFields.Fields.Item("LicTradNum").Value = person.Document;
                    // Set Group RRHH
                    businessObject.GroupCode = 102;
                    
                    
                    // set Branch Code
                    businessObject.BPBranchAssignment.DisabledForBP=SAPbobsCOM.BoYesNoEnum.tNO;
                    businessObject.BPBranchAssignment.BPLID =
                        Int32.Parse(person.GetLastContract().Branches.CodigoSAP);
                    businessObject.BPBranchAssignment.Add();
                    // save new business partner
                    businessObject.Add();
                    // get the new code
                    string newKey = company.GetNewObjectKey();
                    company.GetLastError(out errorCode, out errorMessage);
                    if (errorCode != 0)
                    {
                        log.Success = false;
                        log.ErrorCode = errorCode.ToString();
                        log.ErrorMessage = "SDK: " + errorMessage;
                        _context.SdkErrorLogs.Add(log);
                        _context.SaveChanges();
                        return "ERROR";
                    }
                    else
                    {
                        if (company.InTransaction)
                        {
                            company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            newKey = newKey.Replace("\t1", "");
                            _context.SdkErrorLogs.Add(log);
                            _context.SaveChanges();
                            return newKey;
                        }
                    }
                }
                log.Success = false;
                log.ErrorMessage = "SDK: Not Connected";
                _context.SdkErrorLogs.Add(log);
                _context.SaveChanges();
                return "ERROR";
            }
            catch (Exception ex)
            {
                log.Success = false;
                log.ErrorMessage = "Catch: " + ex.Message;
                _context.SdkErrorLogs.Add(log);
                _context.SaveChanges();
                return "ERROR";
            }
        }

        public string addVoucher(int UserId, Dist_Process process)
        {
            var log = initLog(UserId, BusinessObjectType.Voucher, process.Id.ToString());
            bool approved = false;
            try
            {
                IEnumerable<SapVoucher> dist = _context.Database.SqlQuery<SapVoucher>("SELECT \"ParentKey\",\"LineNum\",\"AccountCode\",sum(\"Debit\") \"Debit\",sum(\"Credit\") \"Credit\", \"ShortName\", null as \"LineMemo\",\"ProjectCode\",\"CostingCode\",\"CostingCode2\",\"CostingCode3\",\"CostingCode4\",\"CostingCode5\",\"BPLId\" " +
                                                                                        " FROM (" +
                                                                                        " select x.\"Id\" \"ParentKey\"," +
                                                                                        "  null \"LineNum\"," +
                                                                                        "  coalesce(b.\"AcctCode\",x.\"CUENTASCONTABLES\") \"AccountCode\"," +
                                                                                        "  CASE WHEN x.\"Indicator\"='D' then x.\"MontoDividido\" else 0 end as \"Debit\"," +
                                                                                        "  CASE WHEN x.\"Indicator\"='H' then x.\"MontoDividido\"else 0 end as \"Credit\"," +
                                                                                        "  x.\"BussinesPartner\" \"ShortName\"," +
                                                                                        "  x.\"Concept\" \"LineMemo\"," +
                                                                                        "  x.\"Project\" \"ProjectCode\"," +
                                                                                        "  f.\"Cod\" \"CostingCode\"," +
                                                                                        "  x.\"PEI\" \"CostingCode2\"," +
                                                                                        "  x.\"PlanEstudios\" \"CostingCode3\"," +
                                                                                        "  x.\"Paralelo\" \"CostingCode4\"," +
                                                                                        "  x.\"Periodo\" \"CostingCode5\"," +
                                                                                        "  x.\"CodigoSAP\" \"BPLId\"" +
                                                                                        " from  (SELECT a.\"Id\",  a.\"Document\",a.\"TipoEmpleado\",a.\"Dependency\",a.\"PEI\"," +
                                                                                        "           a.\"PlanEstudios\",a.\"Paralelo\",a.\"Periodo\",a.\"Project\"," +
                                                                                        "           a.\"Monto\",a.\"Porcentaje\",a.\"MontoDividido\",a.\"segmentoOrigen\",a.\"BussinesPartner\"," +
                                                                                        "           b.\"mes\",b.\"gestion\",e.\"Name\" as Segmento ,d.\"Concept\",d.\"Name\" as CuentasContables,d.\"Indicator\", e.\"CodigoSAP\"" +
                                                                                        "           FROM \"" + CustomSchema.Schema + "\".\"Dist_Cost\" a " +
                                                                                        "               INNER JOIN  \"" + CustomSchema.Schema + "\".\"Dist_Process\" b " +
                                                                                        "               on a.\"DistProcessId\"=b.\"Id\" " +
                                                                                        "           AND a.\"DistProcessId\"= " + process.Id +
                                                                                        "           INNER JOIN  \"" + CustomSchema.Schema + "\".\"Dist_TipoEmpleado\" c " +
                                                                                        "                on a.\"TipoEmpleado\"=c.\"Name\" " +
                                                                                        "           INNER JOIN  \"" + CustomSchema.Schema + "\".\"CuentasContables\" d " +
                                                                                        "              on c.\"GrupoContableId\" = d.\"GrupoContableId\"" +
                                                                                        "           and b.\"BranchesId\" = d.\"BranchesId\" " +
                                                                                        "           and a.\"Columna\" = d.\"Concept\" " +
                                                                                        "           INNER JOIN \"" + CustomSchema.Schema + "\".\"Branches\" e " +
                                                                                        "              on b.\"BranchesId\" = e.\"Id\") x" +
                                                                                        " left join \"" + ConfigurationManager.AppSettings["B1CompanyDB"] + "\".oact b" +
                                                                                        " on x.CUENTASCONTABLES=b.\"FormatCode\"" +
                                                                                        " left join \"" + CustomSchema.Schema + "\".\"Dependency\" d" +
                                                                                        " on x.\"Dependency\"=d.\"Cod\"" +
                                                                                        " left join \"" + CustomSchema.Schema + "\".\"OrganizationalUnit\" f" +
                                                                                        " on d.\"OrganizationalUnitId\"=f.\"Id\"" +
                                                                                        ") V " +
                                                                                        "GROUP BY \"ParentKey\",\"LineNum\",\"AccountCode\", \"ShortName\",\"ProjectCode\",\"CostingCode\",\"CostingCode2\",\"CostingCode3\",\"CostingCode4\",\"CostingCode5\",\"BPLId\";").ToList();
                 var Auxdate = new DateTime(
                    Int32.Parse(process.gestion),
                    Int32.Parse(process.mes),
                    DateTime.DaysInMonth(Int32.Parse(process.gestion), Int32.Parse(process.mes))
                );

                // If process Date is null set last day of the month in proccess
                DateTime date = process.RegisterDate == null ? Auxdate : process.RegisterDate.Value;

                if (company.Connected && dist.Count()>0 && verifyAccounts(UserId,process.Id.ToString(),dist))
                {
                    company.StartTransaction();
                    var dist1 = dist.GroupBy(g => new
                        {
                            g.AccountCode,
                            g.ShortName,
                            g.CostingCode,
                            g.CostingCode2,
                            g.CostingCode3,
                            g.CostingCode4,
                            g.CostingCode5,
                            g.ProjectCode,
                            g.BPLId
                        })
                        .Select(g => new
                        {
                            g.Key.AccountCode,
                            g.Key.ShortName,
                            g.Key.CostingCode,
                            g.Key.CostingCode2,
                            g.Key.CostingCode3,
                            g.Key.CostingCode4,
                            g.Key.CostingCode5,
                            g.Key.ProjectCode,
                            g.Key.BPLId,
                            Credit = g.Sum(s => Double.Parse(s.Credit)),
                            Debit = g.Sum(s => Double.Parse(s.Debit))
                        }).OrderBy(z => z.Debit == 0.00d ? 1 : 0).ThenBy(z => z.AccountCode);

                    if (approved)
                    {
                        SAPbobsCOM.JournalEntries businessObject = (SAPbobsCOM.JournalEntries)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                        // add header Journal Entrie Approved:
                        businessObject.ReferenceDate = date;
                        businessObject.Memo = "Planilla prueba SDK Aprobado " + process.Branches.Abr;
                        businessObject.TaxDate = date;
                        businessObject.Series = Int32.Parse(process.Branches.SerieComprobanteContalbeSAP);
                        businessObject.DueDate = date;

                        // add lines Journal Entrie Approved:
                        businessObject.Lines.SetCurrentLine(0);
                        foreach (var line in dist1)
                        {
                            businessObject.Lines.AccountCode = line.AccountCode;
                            businessObject.Lines.Credit = line.Credit;
                            businessObject.Lines.Debit = line.Debit;
                            if (line.ShortName != null)
                                businessObject.Lines.ShortName = line.ShortName;
                            businessObject.Lines.CostingCode = line.CostingCode;
                            businessObject.Lines.CostingCode2 = line.CostingCode2;
                            businessObject.Lines.CostingCode3 = line.CostingCode3;
                            businessObject.Lines.CostingCode4 = line.CostingCode4;
                            businessObject.Lines.CostingCode5 = line.CostingCode5;
                            businessObject.Lines.ProjectCode = line.ProjectCode;
                            businessObject.Lines.BPLID = Int32.Parse(line.BPLId);
                            businessObject.Lines.Add();
                        }

                        businessObject.Add();

                        string newKey = company.GetNewObjectKey();
                        company.GetLastError(out errorCode, out errorMessage);
                        if (errorCode != 0)
                        {
                            log.Success = false;
                            log.ErrorCode = errorCode.ToString();
                            log.ErrorMessage = "SDK: " + errorMessage;
                            _context.SdkErrorLogs.Add(log);
                            _context.SaveChanges();
                            return "ERROR";
                        }
                        else
                        {
                            if (company.InTransaction)
                            {
                                company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                newKey = newKey.Replace("\t1", "");
                                _context.SdkErrorLogs.Add(log);
                                _context.SaveChanges();
                                return newKey;
                            }
                        }
                    }
                    else
                    {
                        SAPbobsCOM.JournalVouchers businessObject = (SAPbobsCOM.JournalVouchers)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers);
                        // add header Vouvher:

                        businessObject.JournalEntries.ReferenceDate = date;
                        businessObject.JournalEntries.Memo = "Planilla prueba SDK SALOMON " + process.Branches.Abr;
                        businessObject.JournalEntries.TaxDate = date;
                        businessObject.JournalEntries.Series = Int32.Parse(process.Branches.SerieComprobanteContalbeSAP);
                        businessObject.JournalEntries.DueDate = date;

                        // add lines Voucher
                        businessObject.JournalEntries.Lines.SetCurrentLine(0);
                        foreach (var line in dist1)
                        {
                            businessObject.JournalEntries.Lines.AccountCode = line.AccountCode;
                            businessObject.JournalEntries.Lines.Credit = line.Credit;
                            businessObject.JournalEntries.Lines.Debit = line.Debit;
                            if (line.ShortName != null)
                                businessObject.JournalEntries.Lines.ShortName = line.ShortName;
                            businessObject.JournalEntries.Lines.CostingCode = line.CostingCode;
                            businessObject.JournalEntries.Lines.CostingCode2 = line.CostingCode2;
                            businessObject.JournalEntries.Lines.CostingCode3 = line.CostingCode3;
                            businessObject.JournalEntries.Lines.CostingCode4 = line.CostingCode4;
                            businessObject.JournalEntries.Lines.CostingCode5 = line.CostingCode5;
                            businessObject.JournalEntries.Lines.ProjectCode = line.ProjectCode;
                            businessObject.JournalEntries.Lines.BPLID = Int32.Parse(line.BPLId);
                            businessObject.JournalEntries.Lines.Add();
                        }

                        businessObject.Add();

                        string newKey = company.GetNewObjectKey();
                        company.GetLastError(out errorCode, out errorMessage);
                        if (errorCode != 0)
                        {
                            log.Success = false;
                            log.ErrorCode = errorCode.ToString();
                            log.ErrorMessage = "SDK: " + errorMessage;
                            _context.SdkErrorLogs.Add(log);
                            _context.SaveChanges();
                            return "ERROR";
                        }
                        else
                        {
                            if (company.InTransaction)
                            {
                                company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                newKey = newKey.Replace("\t1", "");
                                _context.SdkErrorLogs.Add(log);
                                _context.SaveChanges();
                                return newKey;
                            }
                        }
                    }
                }
                log.Success = false;
                log.ErrorMessage = "SDK: Not Connected or Voucher/Journal Entrie Data Error";
                _context.SdkErrorLogs.Add(log);
                _context.SaveChanges();
                return "ERROR";
            }

            catch (Exception ex)
            {
                log.Success = false;
                log.ErrorMessage = "Catch: " + ex.Message;
                _context.SdkErrorLogs.Add(log);
                _context.SaveChanges();
                return "ERROR";
            }
        }

        public bool verifyAccounts(int UserId, string processId,IEnumerable<SapVoucher> lines)
        {
            foreach (var line in lines)
            {
                if (!checkVoucherAccount(UserId,processId,line,addErrorLog:true))
                    return false;
            }
            return true;
        }

        public bool checkVoucherAccount(int UserId,string processId, SapVoucher line,List<OACT> oact=null, bool addErrorLog=false)
        {
            bool res = true;
            OACT Uoact = null;
            if (oact == null)
            {
                Uoact = _context.Database.SqlQuery<OACT>("select \"AcctCode\",\"AcctName\",\"FormatCode\"," +
                                                             "\"Dim1Relvnt\",\"Dim2Relvnt\",\"Dim3Relvnt\",\"Dim4Relvnt\"," +
                                                             "\"Dim5Relvnt\",\"LocManTran\" from \"" + ConfigurationManager.AppSettings["B1CompanyDB"] + "\".oact" +
                                                             " where \"AcctCode\"='" + line.AccountCode + "';")
                    .FirstOrDefault();
            }
            else
            {
                Uoact = oact.FirstOrDefault(x => x.AcctCode == line.AccountCode);
            }

            if (Uoact == null)
            {
                res = false;
                if (addErrorLog)
                {
                    var log = initLog(UserId, BusinessObjectType.Voucher, processId);
                    log.Success = false;
                    log.ErrorMessage = "AccountCode not Found in SAP: Dist_Cost Id ->" + line.ParentKey;
                    _context.SdkErrorLogs.Add(log);
                    _context.SaveChanges();
                }
            }
            //check Dim1
            if(Uoact.Dim1Relvnt=="Y" && line.CostingCode.IsNullOrWhiteSpace())
            {
                res = false;
                if (addErrorLog)
                {
                    var log = initLog(UserId, BusinessObjectType.Voucher, processId);
                    log.Success = false;
                    log.ErrorMessage = "Dim 1 Required: Dist_Cost Id ->" + line.ParentKey;
                    _context.SdkErrorLogs.Add(log);
                    _context.SaveChanges();
                }
            }
            //check Dim2
            if(Uoact.Dim2Relvnt=="Y" && line.CostingCode2.IsNullOrWhiteSpace())
            {
                res = false;
                if (addErrorLog)
                {
                    var log = initLog(UserId, BusinessObjectType.Voucher, processId);
                    log.Success = false;
                    log.ErrorMessage = "Dim 2 Required: Dist_Cost Id ->" + line.ParentKey;
                    _context.SdkErrorLogs.Add(log);
                    _context.SaveChanges();
                }
            }
            //check Dim3
            if(Uoact.Dim3Relvnt=="Y" && line.CostingCode3.IsNullOrWhiteSpace())
            {
                res = false;
                if (addErrorLog)
                {
                    var log = initLog(UserId, BusinessObjectType.Voucher, processId);
                    log.Success = false;
                    log.ErrorMessage = "Dim 3 Required: Dist_Cost Id ->" + line.ParentKey;
                    _context.SdkErrorLogs.Add(log);
                    _context.SaveChanges();
                }
            }
            //check Dim4
            if(Uoact.Dim4Relvnt=="Y" && line.CostingCode4.IsNullOrWhiteSpace())
            {
                res = false;
                if (addErrorLog)
                {
                    var log = initLog(UserId, BusinessObjectType.Voucher, processId);
                    log.Success = false;
                    log.ErrorMessage = "Dim 4 Required: Dist_Cost Id ->" + line.ParentKey;
                    _context.SdkErrorLogs.Add(log);
                    _context.SaveChanges();
                }
            }
            //check Dim5
            if(Uoact.Dim5Relvnt=="Y" && line.CostingCode5.IsNullOrWhiteSpace())
            {
                res = false;
                if (addErrorLog)
                {
                    var log = initLog(UserId, BusinessObjectType.Voucher, processId);
                    log.Success = false;
                    log.ErrorMessage = "Dim 5 Required: Dist_Cost Id ->" + line.ParentKey;
                    _context.SdkErrorLogs.Add(log);
                    _context.SaveChanges();
                }
            }
            //check Associate Account
            if((Uoact.LocManTran=="Y" && line.ShortName.IsNullOrWhiteSpace())
               || (Uoact.LocManTran == "N" && !line.ShortName.IsNullOrWhiteSpace()))
            {
                res = false;
                if (addErrorLog)
                {
                    var log = initLog(UserId, BusinessObjectType.Voucher, processId);
                    log.Success = false;
                    log.ErrorMessage = "Associate Account: Dist_Cost Id ->" + line.ParentKey;
                    _context.SdkErrorLogs.Add(log);
                    _context.SaveChanges();
                }
            }

            return res;
        }

        public List<dynamic> getBusinessPartners(string col = "CardCode", CustomUser user=null, Branches branch=null)
        {
            List<dynamic> res = new List<dynamic>();
            string[] cols = new string[]
            {
                "\"CardCode\"", "\"CardName\"", "\"LicTradNum\"", "\"CardType\"", "\"GroupCode\"", "\"Series\""
                , "\"Currency\"", "\"City\"", "\"Country\""
            };

            string strcol = "";
            bool first = true;

            foreach (var column in cols)
            {
                strcol += (first ? "" : ",") + "a."+column;
                first = false;
            }

            if (connectedtoHana)
            {
                if (user != null)
                {
                    ADClass auth = new ADClass();
                    var branches = auth.getUserBranches(user);
                    string where = "where ";
                    int f = 0;
                    if (branches == null)
                        where += "false";
                    else
                        foreach (var br in branches)
                        {
                            if (f != 0)
                                where += " or ";
                            where += "b.\"BPLId\" = " + br.CodigoSAP;
                            f++;
                        }

                    string cl = col == "*" ? strcol : "a.\"" + col + "\"";
                    string query = "Select distinct " + cl + " from " + DatabaseName + ".OCRD a " +
                                   "inner join " + DatabaseName + ".CRD8 b " +
                                   " on a.\"CardCode\"=b.\"CardCode\"" + where + " order by a.\"CardCode\"";
                    HanaCommand command = new HanaCommand(query, HanaConn);
                    HanaDataReader dataReader = command.ExecuteReader();
                    if (dataReader.HasRows)
                    {
                        while (dataReader.Read())
                        {
                            if (col == "*")
                            {
                                dynamic x = new JObject();
                                foreach (var column in cols)
                                {
                                    x[column.Replace("\"", "")] = dataReader[column.Replace("\"", "")].ToString();
                                }
                                res.Add(x);
                            }
                            else
                                res.Add(dataReader[col].ToString());
                        }
                    }

                }
                else if (branch != null)
                {
                    ADClass auth = new ADClass();
                    string where = " where b.\"BPLId\" = " + branch.CodigoSAP ;

                    string cl = col == "*" ? strcol : "a.\"" + col + "\"";

                    string query = "Select distinct " + cl + " from " + DatabaseName + ".OCRD a " +
                                   "inner join " + DatabaseName + ".CRD8 b " +
                                   " on a.\"CardCode\"=b.\"CardCode\"" + where;
                    HanaCommand command = new HanaCommand(query, HanaConn);
                    HanaDataReader dataReader = command.ExecuteReader();
                    if (dataReader.HasRows)
                    {
                        while (dataReader.Read())
                        {
                            if (col == "*")
                            {
                                dynamic x = new JObject();
                                foreach (var column in cols)
                                {
                                    x[column.Replace("\"", "")] = dataReader[column.Replace("\"", "")].ToString();
                                }
                                res.Add(x);
                            }
                            else
                                res.Add(dataReader[col].ToString());
                        }
                    }
                }
                else
                {
                    string cl = col == "*" ? col : "\"" + col + "\"";
                    string query = "Select " + cl + " from " + DatabaseName + ".OCRD";
                    HanaCommand command = new HanaCommand(query, HanaConn);
                    HanaDataReader dataReader = command.ExecuteReader();

                    if (dataReader.HasRows)
                    {
                        while (dataReader.Read())
                        {
                            if (col == "*")
                            {
                                dynamic x = new JObject();
                                foreach (var column in cols)
                                {
                                    x[column.Replace("\"", "")] = dataReader[column.Replace("\"", "")].ToString();
                                }
                                res.Add(x);
                            }
                            else
                                res.Add(dataReader[col].ToString());
                        }
                    }
                }
            }

            return res;
        }

        public List<dynamic> getProjects(string col = "PrjCode")
        {
            List<dynamic> res = new List<dynamic>();
            string[] dim1cols = new string[]
            {
                "\"PrjCode\"", "\"PrjName\"", "\"Locked\"", "\"DataSource\"", "\"ValidFrom\"",
                "\"ValidTo\"", "\"Active\"", "\"U_ModalidadProy\"", "\"U_Sucursal\"", "\"U_Tipo\""
            };

            string strcol = "";
            bool first = true;

            foreach (var column in dim1cols)
            {
                strcol += (first ? "" : ",") + column;
                first = false;
            }

            if (connectedtoHana)
            {
                string query = "Select " + strcol + " from " + DatabaseName + ".OPRJ";
                HanaCommand command = new HanaCommand(query, HanaConn);
                HanaDataReader dataReader = command.ExecuteReader();

                if (dataReader.HasRows)
                {
                    while (dataReader.Read())
                    {
                        if (col == "*")
                        {
                            dynamic x = new JObject();
                            foreach (var column in dim1cols)
                            {
                                x[column.Replace("\"", "")] = dataReader[column.Replace("\"", "")].ToString();
                            }
                            res.Add(x);
                        }
                        else
                            res.Add(dataReader[col].ToString());
                    }
                }
            }
            
            return res;
        }
        

        public List<dynamic> getCostCenter(Dimension dimesion,string mes=null,string gestion=null, string col = "PrcCode")
        {
            List<dynamic> res = new List<dynamic>();
            if (connectedtoHana)
            {
                string[][] dim1cols = new string[][]
            {
                new [] {"*"},
                new [] {"\"PrcCode\"", "\"PrcName\"", "\"ValidFrom\"", "\"ValidTo\"", "\"U_TipoUnidadO\""}, 
                new [] {"\"PrcCode\"", "\"PrcName\"", "\"ValidFrom\"", "\"ValidTo\"", "\"U_GestionCC\"", "\"U_AmbitoPEI\"", "\"U_DirectrizPEI\"", "\"U_Indicador\""},
                new [] {"\"PrcCode\"", "\"PrcName\"", "\"ValidFrom\"", "\"ValidTo\"", "\"U_NUM_INT_CAR\"", "\"U_Nivel\""},
                new [] {"\"PrcCode\"", "\"PrcName\"", "\"ValidFrom\"", "\"ValidTo\"", "\"U_PeriodoPARALELO\"", "\"U_Sigla\"", "\"U_Materia\"", "\"U_Paralelo\"", "\"U_ModalidadPARALELO\"", "\"U_EstadoParalelo\"", "\"U_NivelParalelo\"", "\"U_TipoParalelo\""},
                new [] {"\"PrcCode\"", "\"PrcName\"", "\"ValidFrom\"", "\"ValidTo\"", "\"U_GestionPeriodo\"", "\"U_TipoPeriodo\""},
            };

                string strcol = "";
                bool first = true;

                foreach (var column in dim1cols[(int)dimesion])
                {
                    strcol += (first ? "" : ",") + column;
                    first = false;
                }

                string where = (int)dimesion == 0
                    ? ((mes != gestion) ? " where ('2018-02-01 01:00:00' " +
                      "between \"ValidFrom\" and \"ValidTo\")" +
                      "or ('" + gestion + "-" + mes + "-01 01:00:00' > \"ValidFrom\" " +
                      "and \"ValidTo\" is null)" : "")
                    : ((mes != gestion) ? " where \"DimCode\"=" + (int)dimesion +
                      " and (('" + gestion + "-" + mes + "-01 01:00:00' " +
                      "between \"ValidFrom\" and \"ValidTo\")" +
                      "or ('2018-02-01 01:00:00' > \"ValidFrom\" " +
                      "and \"ValidTo\" is null))" : " where \"DimCode\"=" + (int)dimesion);
                string query = "Select " + strcol + " from " + DatabaseName + ".OPRC" + where;
                HanaCommand command = new HanaCommand(query, HanaConn);
                HanaDataReader dataReader = command.ExecuteReader();

                if (dataReader.HasRows)
                {
                    while (dataReader.Read())
                    {
                        if (col == "*")
                        {
                            dynamic x = new JObject();
                            foreach (var column in dim1cols[(int)dimesion])
                            {
                                x[column.Replace("\"", "")] = dataReader[column.Replace("\"", "")].ToString();
                            }
                            res.Add(x);
                        }
                        else
                            res.Add(dataReader[col].ToString());
                    }
                }
            }

            return res;
        }

        public List<object> getParalels()
        {
            List<object> list = new List<object>();
            if (connectedtoHana)
            {
                /*string query = "select a.\"PrcCode\", a.\"U_PeriodoPARALELO\", a.\"U_Sigla\", a.\"U_Paralelo\", b.\"U_CODIGO_DEPARTAMENTO\" "
                + "from ucatolica.oprc a "
                + "inner join admnal.\"T_GEN_PARALELOS\" b "
                +   " on a.\"PrcCode\" = b.\"U_CODIGO_PARALELO\""
                + " WHERE a.\"DimCode\" = " + 4 ;*/

                string query = "select \"PrcCode\", \"U_PeriodoPARALELO\", \"U_Sigla\", \"U_Paralelo\""
                               + "from ucatolica.oprc"
                               + " WHERE \"DimCode\" = " + 4;
                HanaCommand command = new HanaCommand(query, HanaConn);
                HanaDataReader dataReader = command.ExecuteReader();

                if (dataReader.HasRows)
                {
                    while (dataReader.Read())
                    {
                        dynamic o = new JObject();
                        o.cod = dataReader["PrcCode"].ToString();
                        o.periodo = dataReader["U_PeriodoPARALELO"].ToString();
                        o.sigla = dataReader["U_Sigla"].ToString();
                        o.paralelo = dataReader["U_Paralelo"].ToString();
                        //o.OU = dataReader["U_CODIGO_DEPARTAMENTO"].ToString();
                        list.Add(o);
                    }
                }
            }

            return list;
        }

        public string getLastError()
        {
            return errorCode + ": " + errorMessage;
        }

    }
}