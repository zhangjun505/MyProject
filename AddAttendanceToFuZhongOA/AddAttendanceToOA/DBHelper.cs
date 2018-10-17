using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Web.Security;
using System.Security.Cryptography;
namespace AddAttendanceToOA
{
  public static class DBHelper
  {
      static MySqlConnection conn;
      /// <summary>
      /// 打开数据库
      /// </summary>
      /// <param name="conn">要打开的连接对象</param>
      public static void OpenDB(MySqlConnection _conn)
      {
          conn = _conn;
          if (conn.State == ConnectionState.Closed)
              conn.Open();
      }

      /// <summary>
      /// 获取单个值
      /// </summary>
      /// <param name="sql">指定的SQL语句</param>
      /// <returns></returns>
      public static string GetValue(string sql)
      {
          MySqlCommand cmd = new MySqlCommand(sql, conn);
          object o = cmd.ExecuteScalar();
          return o == null ? "" : o.ToString();
      }

      /// <summary>
      /// 执行增删改操作
      /// </summary>
      /// <param name="sql">指定的SQL语句</param>
      /// <returns></returns>
      public static int ExecuteCommand(string sql)
      {
          MySqlCommand cmd = new MySqlCommand(sql, conn);
          try
          {
              return cmd.ExecuteNonQuery();
          }
          catch(Exception ex)
          {
              return 0;
          }
      }

      /// <summary>
      /// 执行存储过程
      /// </summary>
      /// <param name="procName">存储过程名</param>
      /// <param name="paras">参数数组</param>
      public static void ExecuteProc(string procName, params MySqlParameter[] paras)
      {
          MySqlCommand cmd = new MySqlCommand(procName, conn);
          cmd.CommandType = CommandType.StoredProcedure;
          if (paras != null)
          {
              cmd.Parameters.AddRange(paras);
          }
          cmd.ExecuteNonQuery();
      }

      /// <summary>
      /// 获取表中主键Id的最大值
      /// </summary>
      /// <param name="tableName">表名</param>
      /// <param name="primaryKeyId">主键Id名</param>
      /// <returns></returns>
      public static string GetNewId(string tableName, string primaryKeyId)
      {
          string sql = "select max(" + primaryKeyId + ")+1 from " + tableName;
          MySqlCommand cmd = new MySqlCommand(sql, conn);
          object o = cmd.ExecuteScalar();
          return (o == null || o.ToString()=="") ? "1" : o.ToString();
      }

      /// <summary>
      /// 根据不同的加密算法加密字符串
      /// </summary>
      /// <param name="passwordString">要加密的字符串</param>
      /// <param name="passwordFormat">加密算法类型</param>
      /// <returns></returns>
       public static string EncryptPassword(string passwordString,string passwordFormat ) 
       { 
           string  encryptPassword = null;
           if (passwordFormat.ToUpper() == "SHA1")
           {
               encryptPassword = FormsAuthentication.HashPasswordForStoringInConfigFile(passwordString, "SHA1"); 
           }
           else if (passwordFormat.ToUpper() == "MD5") 
           {
               encryptPassword = FormsAuthentication.HashPasswordForStoringInConfigFile(passwordString, "MD5"); 
           }
           else if (passwordFormat.ToUpper() == "DES")
           {
               string key = GenerateKey();

               encryptPassword = EncryptString(passwordString, key);
           }
            return encryptPassword ;
      }

       /// <summary>  
       /// 创建Key  
       /// </summary>  
       /// <returns></returns>  
       static string GenerateKey()
       {
           DESCryptoServiceProvider desCrypto = (DESCryptoServiceProvider)DESCryptoServiceProvider.Create();
           return ASCIIEncoding.ASCII.GetString(desCrypto.Key);
       }

       /// <summary>  
       /// 加密字符串  
       /// </summary>  
       /// <param name="sinputString"></param>  
       /// <param name="Skey"></param>  
       /// <returns></returns>  
       static string EncryptString(string sinputString, string Skey)
       {
           byte[] data = Encoding.UTF8.GetBytes(sinputString);
           DESCryptoServiceProvider DES = new DESCryptoServiceProvider();
           DES.Key = ASCIIEncoding.ASCII.GetBytes(Skey);
           DES.IV = ASCIIEncoding.ASCII.GetBytes(Skey);
           ICryptoTransform desEncrypt = DES.CreateEncryptor();
           byte[] result = desEncrypt.TransformFinalBlock(data, 0, data.Length);
           return BitConverter.ToString(result);
       }  
  }
}
