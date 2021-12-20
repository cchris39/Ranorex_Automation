/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 04/04/2016
 * Time: 12:23 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Linq;
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

namespace RBC_Test
{
	/// <summary>
	/// Description of TestFlow.
	/// </summary>
	[TestModule("51C742FD-3127-4CB1-A429-D9ACA678BDE2", ModuleType.UserCode, 1)]
	public class OrderRequest : ITestModule
	{
		 
		public static RBC_TestRepository repo = RBC_TestRepository.Instance;
		
		static OrderRequest instance = new OrderRequest();
    		
			private string uid = "rbc@nationwideappraisals.com";
    		private string pwd = "Fullmoon#16";
    		private string url = "https://uattest.nas.com/NAS/rbcpvpnas";
		
			private const string streetName = "RBC Sanity test";
			private const string streetType = "Road";
			private const string appFstName = "John";
			private const string appLstName = "Smith";
			private const string transit = "36888";
			private const string language = "E";
			private const string spText = "Test only, please do not proceed. Thanks";
			private const string clientScore = "A";
			private const string contactFstName = "Ms.";
			private const string contactLstName = "Smith";
			//private const string appraisalType1 = "Residential";
			//private const string appraisalType2 = "Full Agricultural";
			private const string builder = "Generic Builder";
			private const string builderCode = "ATXXX";
			private const string siteName = "HQ Tower";
			private const string purchaseType = "Listed on MLS";
						
			private string oriNbr, streetNum, appRefNbr, appPrice, appMtAmount, tel1, tel2, tel3;
			private string nasNbr, serviceType, appraisalType, result;
            private int StrNbr, ClientRef, Price, MortgAmount, ContactNbr1, ContactNbr2, ContactNbr3;
           	private string apprType, ori, loan, proType;
          
           	
		#region Varilables
		string _varCategory = "";
		[TestVariable("5781CA05-BA7C-42A6-A331-F41ECAE32836")]
		public string varCategory
		{
			get { return _varCategory; }
			set { _varCategory = value; }
		}
		
		string _varServiceType = "";
		[TestVariable("C2681C4F-AA4E-4AF7-8432-B3A5EF79AD49")]
		public string varServiceType
		{
			get { return _varServiceType; }
			set { _varServiceType = value; }
		}
		
		
		string _postalCode = "";
		[TestVariable("9F2AF7FC-2EEB-449A-AF55-0AC80B11DD9B")]
		public string postalCode
		{
			get { return _postalCode; }
			set { _postalCode = value; }
		}
		
		string _loanType = "";
		[TestVariable("8B9E9165-0080-4F8C-890B-03764B532F37")]
		public string loanType
		{
			get { return _loanType; }
			set { _loanType = value; }
		}
		
		string _propertyType = "";
		[TestVariable("8F405EE7-49B8-4FBD-B819-1CA18EC6368B")]
		public string propertyType
		{
			get { return _propertyType; }
			set { _propertyType = value; }
		}
		
		string _tenure = "";
		[TestVariable("837439F4-FBE2-4E07-B913-F1D4ECC2FEE3")]
		public string tenure
		{
			get { return _tenure; }
			set { _tenure = value; }
		}
		
		string _mortgageType = "";
		[TestVariable("BBDCC5A2-54D2-411D-BFCB-D6E1DEF9F99C")]
		public string mortgageType
		{
			get { return _mortgageType; }
			set { _mortgageType = value; }
		}
		
		string _programType = "";
		[TestVariable("E6A9F8E3-4995-4030-BFAE-7DC272AEA07C")]
		public string programType
		{
			get { return _programType; }
			set { _programType = value; }
		}
		
		
		string _varApprType = "";
		[TestVariable("F7D8A689-9214-4FBC-A990-E3BBA56DCDE3")]
		public string varApprType
		{
			get { return _varApprType; }
			set { _varApprType = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("BA8DDF7A-67C8-46E3-9CB2-B6EB056762D4")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		#endregion
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public OrderRequest()
		{
			// Do not delete - a parameterless constructor is required!
		}
		
		 public void KillBrowser(string IE)
		{
		}

		public void ClearBrowserCookies(string IE)
		{
		}
		
		public static void GetRandomNbr(int StrNbr, int ClientRef, int Price, int MortgAmount, int ContactNbr1, int ContactNbr2, int ContactNbr3, 
		                                out string streetNum, out string appRefNbr, out string appPrice, out string appMtAmount, 
		                                out string tel1, out string tel2, out string tel3)
        {
            //Convert Integar nunmber to string
            streetNum = Convert.ToString(StrNbr);
            appRefNbr = Convert.ToString(ClientRef);
            appPrice = Convert.ToString(Price);
            appMtAmount = Convert.ToString(MortgAmount);
            tel1 = Convert.ToString(ContactNbr1);
            tel2 = Convert.ToString(ContactNbr2);
            tel3 = Convert.ToString(ContactNbr3);
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
			
			Delay.Milliseconds(300);
			
			var random = new Random();
			int StrNbr = random.Next(1, 900);
			int ClientRef = random.Next(123456789, 987654321);
			int Price = random.Next(500,900)*1000;
			int MortgAmount= random.Next(100,500)*1000;
			int ContactNbr1 = random.Next(111, 999);
			int ContactNbr2 = random.Next(111, 999);
			int ContactNbr3 = random.Next(1111,9999);
		
			//RBC user log in
			Login userLogin = new Login();
			userLogin.login(uid, pwd, url);
			
			//Get Random Number
			GetRandomNbr(StrNbr, ClientRef, Price, MortgAmount, ContactNbr1, ContactNbr2, ContactNbr3, 
			             out streetNum, out appRefNbr, out appPrice, out appMtAmount, out tel1, out tel2, out tel3);
						
			//Request Service
			repo.RBC_Portal.RequestService.Click();
			Delay.Milliseconds(100);
			
			//Instantiate order object
			RequestService orderRequest = new RequestService();
			
			orderRequest.selectRequest(varCategory,varServiceType);
			
			//Get original request number for PI request
			if (varCategory != "Other" && varServiceType == "ProgressInspection")
			{
				string SQL = "SELECT a.`app_request_nbr` FROM app_request a where username regexp 'rbc-mq-qa' and status = 'Completed' and service_residential = 'Full-Service' order by app_request_nbr asc limit 1;";
				DataTable dt = QueryDB.RunQuery(SQL);
				oriNbr = dt.Rows[0][0].ToString().Trim();
				
				//Add on Apr 20, 2016
				string SQL_Cas = "SELECT value FROM app_request_ext a where app_request_ext_type_id = 5 and app_request_nbr  = " + oriNbr + ";";
				DataTable dt_casNbr = QueryDB.RunQuery(SQL_Cas);
				appRefNbr = dt_casNbr.Rows[0][0].ToString().Trim();
				
			}else if (varServiceType == "RuralEstateAgricultural")
			{
				appraisalType = varApprType;
				oriNbr = "";
			}else if (varCategory == "Other" && varServiceType == "ProgressInspection" || varCategory == "Other" && varServiceType == "AppraisalUpdate")
			{
				string SQL1 = "SELECT a.`app_request_nbr` FROM app_request a where username = '" + uid + "' and status = 'Completed' and service_residential = 'Full-Service' order by app_request_nbr desc limit 1;";
				DataTable dt1 = QueryDB.RunQuery(SQL1);
				oriNbr = dt1.Rows[0][0].ToString().Trim();
				//Add on Apr 20, 2016
				string SQL_Cas = "SELECT value FROM app_request_ext a where app_request_ext_type_id = 5 and app_request_nbr  = " + oriNbr + ";";
				DataTable dt_casNbr = QueryDB.RunQuery(SQL_Cas);
				appRefNbr = dt_casNbr.Rows[0][0].ToString().Trim();
			}else
			{  
				appraisalType = "";
				oriNbr = "";
			}
			
			
			//Order request based on cartegory
			if(varCategory == "Linx")
			{
				orderRequest.linxRequest(varServiceType, postalCode, streetNum, streetName, streetType, 
				                         loanType, appPrice, appMtAmount, appFstName, appLstName, appRefNbr, 
				                         transit, language, spText, oriNbr);
			}else if (varCategory == "Homebase" || varCategory == "CASPER")
			{
				orderRequest.portalRequest(varServiceType, postalCode, streetNum, streetName, streetType,
		                          loanType, appPrice, appMtAmount, appFstName, appLstName, appRefNbr, purchaseType,
		                          transit, clientScore, contactFstName, contactLstName,tel1, tel2, tel3, 
		                          language, spText, oriNbr, builder, builderCode, siteName, propertyType, tenure, mortgageType, programType, appraisalType);
				
			}else if (varCategory == "Other")
			{
				orderRequest.portalOtherRequest(varServiceType,oriNbr);
			}
			
			
			//Get Nas nunber and service type based on request type
			if (varServiceType == "Standard" && loanType == "Purchase Price" || varServiceType == "Standard" && loanType == "Refinance Amount")
			{
				nasNbr = orderRequest.GetStdRequestNbr();
				serviceType = orderRequest.GetStdServiceType();
			}else
			{   
				nasNbr = orderRequest.GetOtherRequestNbr();
				serviceType = orderRequest.GetOtherServiceType();
			}
			
			varNasNbr = nasNbr;
			
			Report.Log(ReportLevel.Success, "Success", "Request submit successfully, Nas number is: " + nasNbr);
			Report.Log(ReportLevel.Info, "Information", "Request service type is: " + serviceType);
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(500);
			
		}
	}
}


