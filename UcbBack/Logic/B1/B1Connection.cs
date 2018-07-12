using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using Newtonsoft.Json.Linq;
using Sap.Data.Hana;
using SAPbobsCOM;
using UcbBack.Models;

namespace UcbBack.Logic.B1
{
    public class B1Connection
    {
        private static B1Connection instance;


        private SAPbobsCOM.Company company;
        private HanaConnection HanaConn;
        private int connectionResult;
        private int errorCode = 0;
        private string errorMessage = "";
        private string DatabaseName;
        public enum Dimension
        {
            All,
            OrganizationalUnit,
            PEI,
            PlanAcademico,
            Paralelo,
            Periodo
        }
        public static readonly int OrganizationalUnit = 1;
        public static readonly int PEI = 2;
        public static readonly int PlanAcademico = 3;
        public static readonly int Paralelo = 4;
        public static readonly int Periodo = 5;


        private B1Connection(string db)
        {
            if (TestHanaConection())
            {
                DatabaseName = db;
                string cadenadeconexion = "Server=192.168.18.180:30015;UserID=admnalrrhh;Password=Rrhh12345;Current Schema="+DatabaseName;
                ConnectB1();
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
            get { return instance ?? (instance = new B1Connection("UCATOLICA")); }
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
            company.CompanyDB = "UCATOLICA";
            company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
            company.DbUserName = "RRHH_SDK";
            company.DbPassword = "Rrhh1234";
            company.UserName = "managerrrhh";
            company.Password = "Rrhh1234";
            company.language = SAPbobsCOM.BoSuppLangs.ln_English_Gb;
            company.UseTrusted = true;
            company.LicenseServer = "SAPHANA01:30015";
            company.SLDServer = "SAPHANA01:40000";


            connectionResult = company.Connect();

            if (connectionResult != 0)
            {
                company.GetLastError(out errorCode, out errorMessage);
            }

            return connectionResult;
        }  

        public bool TestHanaConection()
        {
            string cadenadeconexion = "Server=192.168.18.180:30015;UserID=admnalrrhh;Password=Rrhh12345;Current Schema=UCATOLICA";
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


        public void CargaMoneda(out string MonedaLocal, out string MonedaSistema)
        {
            string cadenadeconexion = "Server=192.168.18.180:30015;UserID=admnalrrhh;Password=Rrhh12345;Current Schema=UCATOLICA";
            MonedaLocal = ""; MonedaSistema = "";
            string query = "Select \"MainCurncy\", \"SysCurrncy\" from " + DatabaseName + ".OADM"; //si se quiere llamar a un procedimiento almacenado utilizar "CALL nombreSP"            
            HanaConnection conn = new HanaConnection(cadenadeconexion);
            conn.Open();
            HanaCommand consulta = new HanaCommand(query, conn);
            HanaDataReader resultado = consulta.ExecuteReader();
            if (resultado.HasRows)
            {
                while (resultado.Read())
                {
                    MonedaLocal = resultado["MainCurncy"].ToString();
                    MonedaSistema = resultado["SysCurrncy"].ToString();
                }
            }
            else
            {
                MonedaLocal = "";
                MonedaSistema = "";
            }
            resultado.Close();
            conn.Close();
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
                    businessObject.CardCode = "RC" + person.CUNI;

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

        public List<string> getBusinessPartners(string col = "CardCode")
        {
            string cl = col == "*"?col:"\"" + col + "\"";
            string query = "Select "+cl+" from " + DatabaseName + ".OCRD";
            HanaCommand command = new HanaCommand(query, HanaConn);
            HanaDataReader dataReader = command.ExecuteReader();
            List<string> res = new List<string>();
            if (dataReader.HasRows)
            {
                while (dataReader.Read())
                {
                    res.Add(dataReader[col].ToString());
                }
            }

            return res;
        }

        public List<string> getProjects(string col="PrjCode")
        {
            string query = "Select \"" + col + "\" from " + DatabaseName + ".OPRJ";
            HanaCommand command = new HanaCommand(query, HanaConn);
            HanaDataReader dataReader = command.ExecuteReader();
            List<string> res = new List<string>();
            if (dataReader.HasRows)
            {
                while (dataReader.Read())
                {
                    res.Add(dataReader[col].ToString());
                }
            }

            return res;
        }

        public List<dynamic> getCostCenter(Dimension dimesion,string mes=null,string gestion=null, string col = "PrcCode")
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
                strcol +=  (first? "":",")+column ;
                first = false;
            }

            string where = (int) dimesion == 0
                ? ((mes!=gestion)?" where ('2018-02-01 01:00:00' " +
                  "between \"ValidFrom\" and \"ValidTo\")" +
                  "or ('"+gestion+"-"+mes+"-01 01:00:00' > \"ValidFrom\" " +
                  "and \"ValidTo\" is null)":"")
                : ((mes!=gestion)?" where \"DimCode\"=" + (int) dimesion +
                  " and (('"+gestion+"-"+mes+"-01 01:00:00' " +
                  "between \"ValidFrom\" and \"ValidTo\")" +
                  "or ('2018-02-01 01:00:00' > \"ValidFrom\" " +
                  "and \"ValidTo\" is null))" : " where \"DimCode\"=" + (int)dimesion);
            string query = "Select "+strcol+" from " + DatabaseName + ".OPRC"+where;
            HanaCommand command = new HanaCommand(query, HanaConn);
            HanaDataReader dataReader = command.ExecuteReader();
            List<dynamic> res = new List<dynamic>();
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

            return res;
        }

        public List<object> getParalels()
        {
            string query = "Select \"PrcCode\",\"U_PeriodoPARALELO\",\"U_Sigla\" from " + DatabaseName + ".OPRC ";
            HanaCommand command = new HanaCommand(query, HanaConn);
            HanaDataReader dataReader = command.ExecuteReader();
            List<object> list = new List<object>();
            if (dataReader.HasRows)
            {
                while (dataReader.Read())
                {
                    dynamic o = new JObject();
                    o.cod = dataReader["PrcCode"].ToString();
                    o.periodo = dataReader["U_PeriodoPARALELO"].ToString();
                    o.sigla = dataReader["U_Sigla"].ToString();
                    list.Add(o);
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