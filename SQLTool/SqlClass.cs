using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using System.Data;
using System.Data.SqlClient;
using GlobalClass;

namespace SQLTool
{
    public static class SqlClass
    {
         public static string Conn="";

        public static IEnumerable<T> SelectQry<T>(string Qry,ref string errMsg) where T : new()
        {
            try
            {
                using (IDbConnection cnn = new SqlConnection(Conn))
                {
                    errMsg = "";
                    return cnn.Query<T>(Qry);
                }
                
            }
            catch (Exception ex)
            {
                Log.WriteErrorNlog($"SelectQry=>ExMsg:{ex.Message}=>Qry:{Qry}");
                errMsg =ex.Message.ToString();
                return null;
            }
            
        }
        public static IEnumerable<T> SelectStoredProc<T>(string Qry, DynamicParameters P, ref string errMsg) where T : new()
        {
            try
            {
                using (IDbConnection cnn = new SqlConnection(Conn))
                {
                    errMsg = "";
                    return cnn.Query<T>(Qry, P, commandType: CommandType.StoredProcedure);
                }

            }
            catch (Exception ex)
            {
                Log.WriteErrorNlog($"SelectStoredProc=>ExMsg:{ex.Message}=>Qry:{Qry}");
                errMsg = ex.Message.ToString();
                return null;
            }

        }
        public static int ExecuteCmdQry(string Qry, ref string errMsg)
        {
            try
            {
                using (IDbConnection cnn = new SqlConnection(Conn))
                {
                    return cnn.Execute(Qry);
                }
            }
            catch (Exception ex)
            {
                Log.WriteErrorNlog($"ExecuteCmdQry=>ExMsg:{ex.Message}=>Qry:{Qry}");
                errMsg = ex.Message.ToString();
                return 0;
            }
        }
        public static object ExecutScalarQry(string Qry, ref string errMsg)
        {
            try
            {
                using (IDbConnection cnn = new SqlConnection(Conn))
                {
                    return cnn.ExecuteScalar(Qry);
                }
            }
            catch (Exception ex)
            {
                Log.WriteErrorNlog($"ExecutScalarQry=>ExMsg:{ex.Message}=>Qry:{Qry}");
                errMsg = ex.Message.ToString();
                return null;
            }
        }
    }
}
