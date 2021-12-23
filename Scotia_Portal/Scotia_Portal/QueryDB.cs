/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-05-10
 * Time: 12:08 PM
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
using System.Data;
using MySql.Data.MySqlClient;


using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace Scotia_Portal
{
	/// <summary>
	/// Description of QueryDB.
	/// </summary>
	[TestModule("5B8AF26E-562E-4877-8410-9787DE314C7B", ModuleType.UserCode, 1)]
	
	
	public class QueryDB 
	{
		private MySqlConnection connection; 
		private string server; 
		private string database; 
		private string uid; 
		private string password; 
		private string port;
		
		     	
        static string connectionString = "SERVER=xx.xxx.xx.xxx;DATABASE=appraisal;PORT=3316;UID=anhua;PASSWORD=xxxxxxxx;";
        	
		
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public QueryDB()
		{
			// Do not delete - a parameterless constructor is required!
		}
		
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
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		
}

