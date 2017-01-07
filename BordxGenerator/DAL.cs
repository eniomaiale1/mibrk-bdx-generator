using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using BordxGenerator.Model;
using System.Text.RegularExpressions;


namespace BordxGenerator
{
    class DAL
    {
        public static OleDbConnection conn = new OleDbConnection();

        public static bool CloseConnection()
        {

            try
            {

                conn.Close();
                return true;

            }
            catch (Exception es)
            {
                return false;
            }
        }

        public static bool OpenConnection()
        {

            try
            {
                
                conn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["lloydsDB"].ConnectionString;
                conn.Open();
                return true;

            }
            catch (Exception es)
            {
                return false;
            }

        }

        public static Address GetAddress(string insuredId) {
            Address address = new Address();

            string query = "SELECT Direcciones.Direccion, Direcciones.Ciudad, Direcciones.Estado, Direcciones.[Codigo Potal], [Monedas Paises].Pais " +
                            "FROM Direcciones INNER JOIN[Monedas Paises] ON Direcciones.[Pais ID] = [Monedas Paises].[Pais ID] "+ 
                            "WHERE(((Direcciones.[Asegurado ID])=" + insuredId + ") AND((Direcciones.[Tipo de Direccion ID])=1));";
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
            adapter.Fill(ds);
            dt = ds.Tables[0];
            foreach (DataRow row in dt.Rows)
            {
                address.Line1 = Regex.Replace(row["Direccion"].ToString(), @"\t|\n|\r", "");
                address.City = row["Ciudad"].ToString();
                address.State = row["Estado"].ToString();
                address.Zip = row["Codigo Potal"].ToString();
                address.Country = row["Pais"].ToString();
            }
            return address;

        }

        public static List<ClaimBordx> GetReportData(DateTime from, DateTime to)
        {
            List<ClaimBordx> reportData = null;
            if (DAL.OpenConnection())
            {
                string query = "SELECT Asegurados.[Asegurado ID], Asegurados.Apellidos, Asegurados.Nombres, Asegurados_1.Apellidos as ClaimantLastName, Asegurados_1.Nombres as ClaimantName, " +
                "Status_1.[Numero Poliza], Status_1.[Fecha Desde], Status_1.[Fecha Hasta], Reclamos.[Numero del Reclamo], Worksheets.[Numero del Worksheet], " +
                "Diagnosticos.[ICD 9 Code], Diagnosticos.Diagnostico, Worksheets.[Fecha del Worksheet], [Pagos de Reclamos].[Fecha de Pago], " +
                "Min([Detalle de los Worksheets].[Fecha del Servicio]) AS [MinOfFecha del Servicio], " +
                "Max([Detalle de los Worksheets].[Fecha del Servicio]) AS [MaxOfFecha del Servicio], [Monedas Paises].Pais, Monedas.Simbolo, Worksheets.[Monto Cliente], " +
                "Worksheets.[Monto Cubierto] FROM([Monedas Paises] INNER JOIN ((Diagnosticos INNER JOIN(((Reclamos INNER JOIN Worksheets " +
                "ON Reclamos.[Reclamo ID] = Worksheets.[Reclamo ID]) INNER JOIN(((Asegurados INNER JOIN Status ON Asegurados.[Asegurado ID] = Status.[Asegurado ID]) " +
                "INNER JOIN Status AS Status_1 ON Status.[Status ID Principal] = Status_1.[Status ID]) INNER JOIN Asegurados AS Asegurados_1 " +
                "ON Status_1.[Asegurado ID] = Asegurados_1.[Asegurado ID]) ON Worksheets.[Status ID] = Status.[Status ID]) " +
                "INNER JOIN[Pagos de Reclamos] ON Worksheets.[Pago de Reclamo ID] = [Pagos de Reclamos].[Pago de Reclamo ID]) " +
                "ON Diagnosticos.[Diagnostico ID] = Reclamos.[Diagnostico ID]) INNER JOIN[Detalle de los Worksheets] " +
                "ON Worksheets.[Worksheet ID] = [Detalle de los Worksheets].[Worksheet ID]) ON[Monedas Paises].[Pais ID] = [Detalle de los Worksheets].[Pais Moneda ID]) " +
                "INNER JOIN Monedas ON[Detalle de los Worksheets].[Moneda ID] = Monedas.[Moneda ID] " +
                "WHERE ((([Pagos de Reclamos].[Fecha de Pago]) Between #" + from.ToString("MM/dd/yyyy") + "# And #" + to.ToString("MM/dd/yyyy") + "#)) " +
                "GROUP BY Asegurados.[Asegurado ID], Asegurados.Apellidos, Asegurados.Nombres, " +
                "Asegurados_1.Apellidos, Asegurados_1.Nombres, Status_1.[Numero Poliza], Status_1.[Fecha Desde], Status_1.[Fecha Hasta], Reclamos.[Numero del Reclamo], " +
                "Worksheets.[Numero del Worksheet], Diagnosticos.[ICD 9 Code], Diagnosticos.Diagnostico, Worksheets.[Fecha del Worksheet], " +
                "[Pagos de Reclamos].[Fecha de Pago], [Monedas Paises].Pais, Monedas.Simbolo, [Pagos de Reclamos].[Fecha de Pago], Worksheets.[Monto Cliente], " +
                "Worksheets.[Monto Cubierto] "+
                //"HAVING ((([Pagos de Reclamos].[Fecha de Pago]) Between #" + from.ToString("MM/dd/yyyy") + "# And #" + to.ToString("MM/dd/yyyy") + "#)) "+
                "ORDER BY[Pagos de Reclamos].[Fecha de Pago];";
                //query = "select * from asegurados;";
                reportData = new List<ClaimBordx>();
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
                adapter.Fill(ds);
                dt = ds.Tables[0];
                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        ClaimBordx claim = new ClaimBordx();
                        claim.Address = new Address();
                        claim.InsuredId = row["Asegurado ID"].ToString();
                        claim.Insured = row["Apellidos"].ToString() + " " + row["Nombres"].ToString();
                        claim.Claimant = row["ClaimantLastName"].ToString() + " " + row["ClaimantName"].ToString();
                        claim.PolicyNumber = row["Numero Poliza"].ToString();

                        claim.EffectiveDate = DateTime.Parse(row["Fecha Desde"].ToString());
                        claim.ExpirationDate = DateTime.Parse(row["Fecha Hasta"].ToString());
                        //string claimnumber = row["Numero del Reclamo"].ToString();

                        int claimnro = int.Parse(row["Numero del Reclamo"].ToString());

                        claim.ClaimNumber = "L" + claimnro.ToString("0000000") + "/" + row["Numero del Worksheet"].ToString();
                        claim.Year = claim.ExpirationDate.Year;
                        claim.LossDateFrom = DateTime.Parse(row["MinOfFecha del Servicio"].ToString());
                        claim.LossDateTo = DateTime.Parse(row["MaxOfFecha del Servicio"].ToString());
                        string icd9 = row["ICD 9 Code"].ToString();
                        claim.LossDescription = string.IsNullOrEmpty(icd9) ? row["Diagnostico"].ToString() : icd9 + "-" + row["Diagnostico"].ToString();
                        claim.LossLocation = row["Pais"].ToString();
                        claim.OriginalCurrency = row["Simbolo"].ToString();
                        claim.AmountClaimed = double.Parse(row["Monto Cliente"].ToString());
                        claim.AmountPaid = double.Parse(row["Monto Cubierto"].ToString());
                        claim.DateClaimMade = DateTime.Parse(row["Fecha del Worksheet"].ToString());
                        claim.DateClaimPaid = DateTime.Parse(row["Fecha de Pago"].ToString());
                        claim.DateClaimNotified = DateTime.Parse(row["Fecha del Worksheet"].ToString());
                        reportData.Add(claim);
                    }
                    catch (Exception es) { }
                }

                foreach (ClaimBordx claim in reportData) {
                    claim.Address = GetAddress(claim.InsuredId);
                }
                DAL.CloseConnection();
            }
            return reportData;
        }

        public static List<ClaimBordx> GetFees(DateTime from, DateTime to)
        {
            List<ClaimBordx> reportData = null;
            if (DAL.OpenConnection())
            {
                string query = "SELECT [Checks Generated].[Payment Date],  [Checks Generated].[Pay to], [Checks Generated].Amount, [Checks Generated].[Check Nro], " +
                                " [Checks Generated].[Month Coverage], [Checks Generated].Memo FROM [Checks Generated] " +
                                " WHERE [Checks Generated].[Payment Date] Between #" + from.ToString("MM/dd/yyyy") + "# And #" + to.ToString("MM/dd/yyyy") + "# " +
                                " AND Year([Checks Generated].[Payment Date]) = " + to.ToString("yyyy") + 
                                " ORDER BY [Checks Generated].[Payment Date] ";

                reportData = new List<ClaimBordx>();
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
                adapter.Fill(ds);
                dt = ds.Tables[0];
                foreach (DataRow row in dt.Rows)
                {
                    ClaimBordx claim = new ClaimBordx();
                    claim.Address = new Address();
                    claim.DateFeesPaid = DateTime.Parse(row["Payment Date"].ToString());
                    claim.FeesPaid = double.Parse(row["Amount"].ToString());
                    claim.LossDescription = row["Pay to"].ToString() + " (Check #: " + row["Check Nro"].ToString() + ")";

                    reportData.Add(claim);
                }
                DAL.CloseConnection();
            }
            return reportData;
        }

        public static List<Period> GetPeriods() {
            List<Period> result = null;
            
            if (DAL.OpenConnection()) { 
                result = new List<Period>();

                string query = "SELECT [Periods of Coverage].ID, [Periods of Coverage].[Fecha Desde], [Periods of Coverage].[Fecha Hasta], [Periods of Coverage].Contract, " +
                    "[Periods of Coverage].Reference FROM [Periods of Coverage] ORDER BY[Periods of Coverage].[Fecha Hasta];";
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
                adapter.Fill(ds);
                dt = ds.Tables[0];
                foreach (DataRow row in dt.Rows)
                {
                    Period period = new Period();
                    period.id = int.Parse(row["ID"].ToString());
                    period.From = DateTime.Parse(row["Fecha Desde"].ToString());
                    period.To = DateTime.Parse(row["Fecha Hasta"].ToString());
                    period.Contract = row["Contract"].ToString();
                    period.Reference = row["Reference"].ToString();
                    result.Add(period);
                }
                DAL.CloseConnection();
            }
            return result;
        }

    }
}
