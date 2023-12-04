using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ScriptsExecutionUtility.Models
{
    public class DataBaseConnectivity
    {
        string ConnectionString = ConfigurationManager.AppSettings["Bill800ConnectionString"];
        public List<SqlParameter> returnSppram(Object obj)
        {
            List<SqlParameter> parameters = new List<SqlParameter>();

            Type objectType = obj.GetType();
            PropertyInfo[] properties = objectType.GetProperties();

            foreach (PropertyInfo property in properties)
            {
                parameters.Add(new System.Data.SqlClient.SqlParameter("@" + property.Name, property.GetValue(obj)));
            }
            return parameters;

        }
        public List<Object> ConverttoObject(DataTable data, Type objectType)
        {

            List<Object> objectList = new List<Object>();

            foreach (DataRow row in data.Rows)
            {
                object obj = Activator.CreateInstance(objectType);

                foreach (DataColumn column in data.Columns)
                {
                    string propertyName = column.ColumnName;
                    object propertyValue = row[column];

                    PropertyInfo property = objectType.GetProperty(propertyName);
                    if (property != null && property.CanWrite)
                    {
                        property.SetValue(obj, Convert.ToString(propertyValue));
                    }
                }

                objectList.Add(obj);
            }

            return objectList;
        }
        public DataTable ExecuteProc(string pProcedureName, List<SqlParameter> param)
        {
            DataTable Dbres = new DataTable();
            using (SqlConnection sqlconobj = new SqlConnection(ConnectionString))
            {
                using (SqlCommand sqlcmdobj = new SqlCommand(pProcedureName, sqlconobj))
                {
                    sqlcmdobj.CommandType = CommandType.StoredProcedure;
                    sqlcmdobj.CommandTimeout = 300000;
                    //mobjCommand.CommandTimeout = CommandTimeOut;// 20;
                    if (param != null)
                    {
                        for (int i = 0; i < param.Count; i++)
                        {
                            if (param[i] != null)
                            {
                                sqlcmdobj.Parameters.Add(param[i]);
                            }
                        }
                    }
                    try
                    {
                        sqlconobj.Open();
                        SqlDataAdapter sqladpobj;
                        sqladpobj = new SqlDataAdapter(sqlcmdobj);
                        DataSet mDataSet = new DataSet();
                        sqladpobj.Fill(mDataSet);
                        Dbres = mDataSet.Tables[0];


                    }
                 
                    catch (Exception ex)
                    {
                     return null;
                    }
                    finally
                    {
                        sqlcmdobj.Dispose();
                        sqlconobj.Close();
                        sqlconobj.Dispose();
                    }
                    return Dbres;
                }
            }
        }
    }
}
