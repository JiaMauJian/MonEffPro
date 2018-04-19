using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ClosedXML
{
    public static class NoOraConn
    {
        const string conn = @"USER ID=edauser;PASSWORD=edauser;DATA SOURCE=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=mmsreport.motech.corp)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=mmsrep)))";

        public static DataTable Query(string sql)
        {
            DataTable dt = new DataTable();

            using(OracleConnection connection = new OracleConnection(conn))
            {
                try
                {
                    connection.Open();
                    using (OracleDataAdapter daORA = new OracleDataAdapter(sql, connection))
                    {
                        daORA.Fill(dt);

                        /*foreach (DataRow r in dt.Rows)
                        {
                            for (int i = 0; i<= dt.Columns.Count - 1; i++)
                            {
                                if (Convert.IsDBNull(r.ItemArray[i]))
                                {
                                    try
                                    {
                                        if (dt.Columns[i].DataType.ToString() == "System.Decimal")
                                        {
                                            r.ItemArray[i] = 0;
                                        }
                                        else if (dt.Columns[i].DataType.ToString() == "System.String")
                                        {
                                            r.ItemArray[i] = "";
                                        }
                                        else if (dt.Columns[i].DataType.ToString() == "System.DateTime")
                                        {

                                        }
                                        else
                                        {
                                            r.ItemArray[i] = "";
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        r.ItemArray[i] = "";
                                    }

                                }

                            }

                        }*/

                    }
                }
                catch (OracleException ex)
                {
                    MessageBox.Show(ex.Message);                    
                }
                finally
                {
                    connection.Close();
                }
            }
            return dt;
        }
    }
}
