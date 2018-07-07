using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Status_changer
{
    class DBContext
    {
        static public DataTable GetConsStatus()
        {
            DataTable result = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(Properties.Settings.Default.conString))
                {
                    con.Open();

                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        cmd.CommandText = "select id, Consignment, Date, Code, Commentary, EventDepot from InvoicesStatuses WHERE InMainframe = 0;";

                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                result.Load(reader);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return result;
            }
            return result;
        }

        static public void ChangeRecordStatus(int id)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(Properties.Settings.Default.conString))
                {
                    con.Open();

                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        cmd.CommandText = "Update InvoicesStatuses Set InMainframe = 1 Where id = @id";

                        cmd.Parameters.Clear();

                        cmd.Parameters.Add("@id", SqlDbType.Int);
                        cmd.Parameters["@id"].Value = id;

                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }
    }

}