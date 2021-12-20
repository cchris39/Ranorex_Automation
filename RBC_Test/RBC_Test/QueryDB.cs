/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 02/04/2016
 * Time: 2:26 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Data;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace RBC_Test
{
	/// <summary>
	/// Description of DatabaseQuery.
	/// </summary>
	[TestModule("BE4FB7B1-42DB-4F31-A506-36C57D48FD15", ModuleType.UserCode, 1)]
	public class QueryDB
	{
		private MySqlConnection connection; 
		private string server; 
		private string database; 
		private string uid; 
		private string password; 
		private string port;
		
		     	
        static string connectionString = "SERVER=xx.xxx.xx.xxx;DATABASE=appraisal;PORT=3316;UID=anhua;PASSWORD=xxxxxx;";
        	
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public QueryDB()
		{
			// Do not delete - a parameterless constructor is required!
		}

		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		
		//open connection to database 
		private bool OpenConnection() 
		{ 
			try 
			{ connection.Open(); 
			  return true;
			} 
			catch (MySqlException ex)
			{ 
				//When handling errors, you can your application's response based 
				//on the error number. 
				//The two most common error numbers when connecting are as follows: 
				//0: Cannot connect to server. 
				//1045: Invalid user name and/or password. 
				switch (ex.Number)
				{ 
					case 0:
					MessageBox.Show("Cannot connect to server. Contact administrator"); 
					break; 
					
					case 1045: 
					MessageBox.Show("Invalid username/password, please try again"); 
					break; 
				} 
				return false; 
			} 
		}
		
		//Close connection 
		private bool CloseConnection() 
		{ 
			try 
			{ connection.Close();
			  return true;
			} 
			catch (MySqlException ex) 
			{ MessageBox.Show(ex.Message); 
			  return false;
			}
		}
		
		 public static DataTable RunQuery(string SQL)
      {
         
         MySqlConnection connection = new MySqlConnection(connectionString);
         MySqlDataReader reader = null;
         DataTable dt = new DataTable();
         try{
            MySqlCommand command = connection.CreateCommand();
            command.CommandText = SQL;
            connection.Open();
            reader = command.ExecuteReader();
            
            //load the mysql reader into a databable - this allows the connection to the db to be closed while still
            //having access to the data that the query returned
            dt.Load(reader);
         }catch(Exception e){
            // Console.WriteLine("Exception is MySQLConnector::RunQuery - " + e.ToString());
            
                MessageBox.Show("Exception is MySQLConnector::RunQuery -- " + e.ToString() );
            
         }finally
         {
            reader.Close();
            connection.Close();
         }
         return dt;
      }
      
      /// <summary>
      /// Returns the count of records found for a count query.  Example SELECT COUNT(*) FROM hdlogin.
      /// Use this when you just was a count and don't need any data returned.
      /// </summary>
      /// <param name="SQL"></param>
      /// <returns></returns>
      public static int CountQuery(string SQL)
      {
         MySqlConnection connection = new MySqlConnection(connectionString);
         int count=0;
         try{
            MySqlCommand command = connection.CreateCommand();
            command.CommandText = SQL;
            connection.Open();
            count = Convert.ToInt32(command.ExecuteScalar());
         }
         catch (Exception e){
            // Console.WriteLine("Exception is MySQLConnector::CountQuery - " + e.ToString());
            
            MessageBox.Show("Cannot connect to server.  Contact administrator -- " + e.ToString() );
            
         }
         finally{
            connection.Close();
         }
         return count;
      }
      
      /// <summary>
      /// Returns the number of rows affected by the supplied query.  Use this if you need to run an UPDATE or INSERT query.
      /// </summary>
      /// <param name="SQL"></param>
      /// <returns></returns>
      public static int UpdateQuery(string SQL)
      {
         MySqlConnection connection = new MySqlConnection(connectionString);
         int rowsAffected=0;
         try{
            MySqlCommand command = connection.CreateCommand();
            command.CommandText = SQL;
            connection.Open();
            rowsAffected = Convert.ToInt32(command.ExecuteNonQuery());
         }
         catch (Exception e){
            // Console.WriteLine("Exception is MySQLConnector::UpdateQuery - " + e.ToString());
            MessageBox.Show("Exception is MySQLConnector::UpdateQuery -- " + e.ToString() );
            
         }
         finally{
            connection.Close();
         }
         return rowsAffected;
      }
      
		
	}
	
}
