using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using Newtonsoft.Json.Linq;
using Sap.Data.Hana;
using SAPbobsCOM;
using UcbBack.Models;
using UcbBack.Models.Not_Mapped;

namespace UcbBack.Logic.B1
{
    public class B1Connection
    {
        private static B1Connection instance=null;


        private SAPbobsCOM.Company company;
        private HanaConnection HanaConn;
        private int connectionResult;
        private int errorCode = 0;
        private string errorMessage = "";
        private string DatabaseName;
        public bool connectedtoHana=false;

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
            connectedtoHana = TestHanaConection();
            if (connectedtoHana)
            {
                DatabaseName = ConfigurationManager.AppSettings["B1CompanyDB"];
                //string cadenadeconexion = "Server=192.168.18.180:30015;UserID=admnalrrhh;Password=Rrhh12345;Current Schema="+DatabaseName;
                //string cadenadeconexion = "Server=SAPHANA01:30015;UserID=SDKRRHH;Password=Rrhh1234;Current Schema=UCBTEST"+DatabaseName;
                string cadenadeconexion = "Server=" + ConfigurationManager.AppSettings["B1Server"] +
                                          ";UserID=" + ConfigurationManager.AppSettings["HanaBDUser"] +
                                            ";Password=" + ConfigurationManager.AppSettings["HanaPassword"] +
                                            ";Current Schema=" + ConfigurationManager.AppSettings["HanaBD"];
                //ConnectB1();
                HanaConn = new HanaConnection(cadenadeconexion);
                HanaConn.Open();
            }
            else
            {
                instance = null;
            }
        }

        public static B1Connection Instance
        {
            // get { return instance ?? (instance = new B1Connection("UCATOLICA")); }
            get { return instance ?? (instance = new B1Connection()); }
        }


        public bool DisconnectB1()
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

        public int ConnectB1()
        {
            company = new SAPbobsCOM.Company();

            company.Server = "SAPHANA01:30015";
            company.CompanyDB = "UCBTEST";
            company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
            company.DbUserName = "DESARROLLO1";
            company.DbPassword = "Rrhh12345";
            company.UserName = "manager7";
            company.Password = "sandra2018";
            company.language = SAPbobsCOM.BoSuppLangs.ln_English_Gb;
            company.UseTrusted = true;
            company.LicenseServer = "SAPHANA01:30015";
            company.SLDServer = "SAPHANA01:40000";

            /*company.Server = ConfigurationManager.AppSettings["B1Server"];
            company.CompanyDB = ConfigurationManager.AppSettings["B1CompanyDB"];
            company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
            company.DbUserName = ConfigurationManager.AppSettings["B1DbUserName"];
            company.DbPassword = ConfigurationManager.AppSettings["B1DbPassword"];
            company.UserName = ConfigurationManager.AppSettings["B1UserName"];
            company.Password = ConfigurationManager.AppSettings["B1Password"];
            company.language = SAPbobsCOM.BoSuppLangs.ln_English_Gb;
            company.UseTrusted = true;
            company.LicenseServer = ConfigurationManager.AppSettings["B1LicenseServer"];
            company.SLDServer = ConfigurationManager.AppSettings["B1SLDServer"];*/



            connectionResult = company.Connect();
            var x = company.Connected;
        if (connectionResult != 0)
            {
                company.GetLastError(out errorCode, out errorMessage);
            }
            var _context = new ApplicationDbContext();
//            var yo = _context.Person.FirstOrDefault(xx => xx.CUNI == "RFA940908");
            var yo = new People();
            yo.CUNI = "QWE123456";
            yo.Document = "1234567";
            yo.FirstSurName = "Valdes";
            yo.FirstSurName = "Juan";
            if (yo == null)
                return -1;
            var  res = personToBP(yo);
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

        public string addPersonToB1(People person)
        {
            string message="";
            try
            {
                if (company.Connected)
                {
                    company.StartTransaction();
                    SAPbobsCOM.EmployeesInfo oEmployeesInfo = (SAPbobsCOM.EmployeesInfo)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo);

                    oEmployeesInfo.FirstName = person.Names;
                    oEmployeesInfo.LastName = person.FirstSurName;
                    oEmployeesInfo.Gender = person.Gender == "M" ? BoGenderTypes.gt_Male : BoGenderTypes.gt_Female;
                    oEmployeesInfo.DateOfBirth = person.BirthDate;

                    oEmployeesInfo.Add();
                    string newKey = company.GetNewObjectKey();
                    company.GetLastError(out errorCode, out errorMessage);
                    if (errorCode != 0)
                    {
                        message = "Error - "+errorCode+": "+errorMessage;
                    }
                    else
                    {
                        if (company.InTransaction)
                        {
                            company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            newKey = newKey.Replace("\t1", "");
                            message = newKey + "- successful!";
                        }
                    }  
                }
            }
            catch (Exception ex)
            {
                message = message + " - Error: " + ex.Message;
                DisconnectB1();
            }
            return message;     
        }

        public string personToBP(People person)
        {
            string message = "";
            try
            {
                if (company.Connected)
                {
                    company.StartTransaction();
                    SAPbobsCOM.BusinessPartners businessObject = (SAPbobsCOM.BusinessPartners)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);

                    businessObject.CardName = person.FirstSurName + " " + person.Names;
                    businessObject.CardType = SAPbobsCOM.BoCardTypes.cCustomer;
                    businessObject.CardCode = "R" + "123456";
                    var lis = businessObject.UserFields.Fields;
                    businessObject.UserFields.Fields.Item("").Value = 1234567;

                    businessObject.Add();
                    string newKey = company.GetNewObjectKey();
                    company.GetLastError(out errorCode, out errorMessage);
                    if (errorCode != 0)
                    {
                        message = "Error - " + errorCode + ": " + errorMessage;
                    }
                    else
                    {
                        if (company.InTransaction)
                        {
                            company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            newKey = newKey.Replace("\t1", "");
                            message = newKey + "- successful!";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                message = message + " - Error: " + ex.Message;
                DisconnectB1();
            }
            return message;     
        }

        public string addVoucher(IEnumerable<SapVoucher> lines, Dist_Process process)
        {
            string message = "";
            try
            {
                if (company.Connected)
                {
                    company.StartTransaction();
                    SAPbobsCOM.JournalVouchers businessObject = (SAPbobsCOM.JournalVouchers)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers);

                    
                    string newKey = company.GetNewObjectKey();
                    company.GetLastError(out errorCode, out errorMessage);
                    if (errorCode != 0)
                    {
                        message = "Error - " + errorCode + ": " + errorMessage;
                    }
                    else
                    {
                        if (company.InTransaction)
                        {
                            company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            newKey = newKey.Replace("\t1", "");
                            message = newKey + "- successful!";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                message = message + " - Error: " + ex.Message;
                DisconnectB1();
            }
            return message;
        }

        public List<string> getBusinessPartners(string col = "CardCode")
        {
            List<string> res = new List<string>();
            if (connectedtoHana)
            {
                string cl = col == "*" ? col : "\"" + col + "\"";
                string query = "Select " + cl + " from " + DatabaseName + ".OCRD";
                HanaCommand command = new HanaCommand(query, HanaConn);
                HanaDataReader dataReader = command.ExecuteReader();

                if (dataReader.HasRows)
                {
                    while (dataReader.Read())
                    {
                        res.Add(dataReader[col].ToString());
                    }
                }
            }

            return res;
        }

        public List<string> getProjects(string col="PrjCode")
        {
            List<string> res = new List<string>();
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
                string query = "Select \"PrcCode\",\"U_PeriodoPARALELO\",\"U_Sigla\",\"U_Paralelo\" from " + DatabaseName + ".OPRC ";
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