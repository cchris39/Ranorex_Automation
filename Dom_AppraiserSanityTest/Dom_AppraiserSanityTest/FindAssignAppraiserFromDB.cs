/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 13/01/2016
 * Time: 10:36 AM
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
//Add MySql Library
using MySql.Data.MySqlClient;
using System.Data;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace Dom_AppraiserSanityTest
{
	/// <summary>
	/// Description of FindAssignAppraiserFromDB.
	/// Scenario: Get the assign appraiser from Database based on the client submit request
	/// </summary>
	[TestModule("42BD977B-0843-4FA7-9772-DE93AC80FB8A", ModuleType.UserCode, 1)]
	public class FindAssignAppraiserFromDB : ITestModule
	{
		private MySqlConnection connection;
    	private string server;
    	private string database;
    	private string uid;
    	private string password;
    	private string port;
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public FindAssignAppraiserFromDB()
		{
			// Do not delete - a parameterless constructor is required!
			//Initialize();
		}
		
	#region Variables

	string _varNasNbr = "";
	[TestVariable("9AA3A688-95FA-4916-8F15-01DE2899AD31")]
	public string varNasNbr
	{
		get { return _varNasNbr; }
		set { _varNasNbr = value; }
	}
	
	#endregion
		
		string _varAppraiser = "";
		[TestVariable("E3B4127E-13B9-4122-8399-29FFA08DEBBD")]
		public string varAppraiser
		{
			get { return _varAppraiser; }
			set { _varAppraiser = value; }
		}
		
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
		
			//Construct a connection string
			server = "xx.xxx.xx.xxx";    
        	database = "appraisal";
        	uid = "anhua";
        	password = "xxxxxxxx";
        	port = "3316";      //3316
        	string connectionString;
        	connectionString = "SERVER=" + server + ";" + "DATABASE=" + 
			database + ";" + "PORT=" + port + ";" + "UID=" + uid + ";" + "PASSWORD=" + password + ";";
		
        		try
    			{
        		//open connection to database
    			connection = new MySqlConnection(connectionString);
    			connection.Open();
        		
        		
        			//Select statement
    				string query = "SELECT a.appraiser FROM Appraisal.app_Request a where a. app_request_nbr = " + varNasNbr;
    				
    				//string query2 = "SELECT u.`username`, u.`first_name`, u.`last_name`, u.`province`, d.expiry_date FROM user_info u join appraiser_document d on d.username = u.username +
									//where u.username = " + varAppraiser + " order by d.expiry_date desc limit 1;";
    			
        			//Create Command
       				 MySqlCommand cmd = new MySqlCommand(query, connection);
       				 //Create a data reader and Execute the command
        			MySqlDataReader dataReader = cmd.ExecuteReader();
        
        			//Read the data and store them in the list
        			while (dataReader.Read())
        			{
        			varAppraiser = dataReader[0].ToString();
        			//MessageBox.Show(varAppraiser);
        			}
					
        			//close Data Reader
        			dataReader.Close();

        			//close Connection
        			connection.Close();;
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
                	MessageBox.Show("Cannot connect to server.  Contact administrator");
                	break;

            		case 1045:
                	MessageBox.Show("Invalid username/password, please try again");
                	break;
        			}
        		}
    				
	    	}
 	}
}
