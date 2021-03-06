/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-04-25
 * Time: 3:33 PM
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

namespace RBC_Test
{
	/// <summary>
	/// Description of SearchRequest.
	/// </summary>
	[TestModule("EB12BFC3-D205-4BDD-A878-6E5000E8FFEF", ModuleType.UserCode, 1)]
	public class SearchRequest : ITestModule
	{
		public static RBC_TestRepository repo = RBC_TestRepository.Instance;
		
		static SearchRequest instance = new SearchRequest();

		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public SearchRequest()
		{
			// Do not delete - a parameterless constructor is required!
		}
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("C65160ED-C31D-4639-8C67-842C0F0878B5")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varDate = "";
		[TestVariable("089C9041-A61C-494D-9DCB-21EF658A4303")]
		public string varDate
		{
			get { return _varDate; }
			set { _varDate = value; }
		}
		
		string _varAddress = "";
		[TestVariable("CA0F8E79-AC75-4851-8A5A-501F41D5D1A6")]
		public string varAddress
		{
			get { return _varAddress; }
			set { _varAddress = value; }
		}
		
		string _varAppName = "";
		[TestVariable("912649DD-F7BA-4A3E-AF69-328F83C2B552")]
		public string varAppName
		{
			get { return _varAppName; }
			set { _varAppName = value; }
		}
		
		string _varCasperNbr = "";
		[TestVariable("A07FB221-446E-4475-99D2-DAC7CF965274")]
		public string varCasperNbr
		{
			get { return _varCasperNbr; }
			set { _varCasperNbr = value; }
		}
		
		string _varPoCode = "";
		[TestVariable("CFD3C097-1354-41E7-AEC7-DF3A4DB10A2F")]
		public string varPoCode
		{
			get { return _varPoCode; }
			set { _varPoCode = value; }
		}
		
		string _varStatus = "";
		[TestVariable("1CE797A2-5851-4FDE-8881-B6CB73072FD8")]
		public string varStatus
		{
			get { return _varStatus; }
			set { _varStatus = value; }
		}
		
		string _varUid = "";
		[TestVariable("5B7FB956-CBD0-4AA0-BDF0-C10958AB65F0")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("E5D6F1B9-D3B8-420E-A7B3-0E6FC302731C")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varUrl = "";
		[TestVariable("9A5B3AE2-6DD2-4F7E-9F5A-DAA9FAFD0AC4")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		#endregion
		
		//private string uid = "rbctest@nas.com";
    	//private string pwd = "xxxxxxx";
    	//private string url = "https://uattest.nas.com/NAS/rbcpvpnas";
    	
    	private string status, regDate, appName, address, city, province, serviceType, casperNbr;
    	
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		public void findByNasNbr(string nbr)
		{
			repo.RBC_Portal.SearchFilter.Click();
			Delay.Milliseconds(100);
			
			//Search BY Nas Number
			repo.RBC_Portal.SearchRequest.NasReqNum.PressKeys(nbr);
			repo.RBC_Portal.SearchRequest.ViewGroupReq.Click();
			repo.RBC_Portal.SearchRequest.SearchSubmitBtn.Click();
			Delay.Milliseconds(100);
			
			string nasNbrFound = repo.RBC_Portal.SearchRequest.NasNumberTag.InnerText.Trim();
			
			Report.Log(ReportLevel.Success, "Success", "Request " + varNasNbr + " found.");
			Validate.AreEqual(nasNbrFound, nbr);
			
		}
		
		public void findByDate(string dateFrom, string dateTo)
		{
			//Search BY Date Range
			Delay.Milliseconds(100);
			repo.RBC_Portal.SearchRequest.CreateDateFrom.TagValue = dateFrom;
			repo.RBC_Portal.SearchRequest.CreateDateTo.TagValue = dateTo;
			repo.RBC_Portal.SearchRequest.ViewGroupReq.Click();
			repo.RBC_Portal.SearchRequest.SearchSubmitBtn.Click();
			Delay.Milliseconds(100);
			
			string nasNbrFound = repo.RBC_Portal.SearchRequest.NasNumberTag.InnerText.Trim();
			
			Report.Log(ReportLevel.Success, "Success", "Request " + varNasNbr + " found.");
			Validate.AreEqual(nasNbrFound, varNasNbr);
		}
		
		public void findByCasperNbr(string Nbr)
		{
			repo.RBC_Portal.SearchFilter.Click();
			Delay.Milliseconds(100);
			
			//Search BY Nas Number
			repo.RBC_Portal.SearchRequest.CASPER_LinxNbr.PressKeys(varCasperNbr);
			repo.RBC_Portal.SearchRequest.ViewGroupReq.Click();
			repo.RBC_Portal.SearchRequest.SearchSubmitBtn.Click();
			Delay.Milliseconds(100);
			
			string nasNbrFound = repo.RBC_Portal.SearchRequest.NasNumberTag.InnerText.Trim();
			
			Report.Log(ReportLevel.Info, "Searching request by CASPER/Linx number: " + varCasperNbr);
			Report.Log(ReportLevel.Success, "Success", "Request " + varNasNbr + " found.");
			Validate.AreEqual(nasNbrFound, varNasNbr);
		}
		
		
		
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			//RBC user log in
			//Login userLogin = new Login();
			//userLogin.login(varUid, varPwd,varUrl);
			
			
			//GET CURRENT DATE VALUE
			var curDate = System.DateTime.Today.ToString("yyyy-MM-dd");
						
			Delay.Milliseconds(100);
						
			//PERFORM SEARCH
			//Search By Nas Number
			findByNasNbr(varNasNbr);
			
			//Search by Address
			Report.Log(ReportLevel.Info, "Information", "Searching request " + varNasNbr + " by address: " + varAddress);
			repo.RBC_Portal.SearchFilter.Click();
			Delay.Milliseconds(100);
			repo.RBC_Portal.SearchRequest.Address.PressKeys(varAddress);
			findByDate(varDate, curDate);
						
			Delay.Milliseconds(100);
			//Search by Postal Code
			Report.Log(ReportLevel.Info, "Information", "Searching request " + varNasNbr + " by postal code: " + varPoCode);
			repo.RBC_Portal.SearchFilter.Click();
			Delay.Milliseconds(100);
			repo.RBC_Portal.SearchRequest.Postalcode.PressKeys(varPoCode);
			findByDate(varDate, curDate);
			
			
			//Search by CASPER number
			findByCasperNbr(varCasperNbr);
			
			//Close Browser
			//Host.Local.KillBrowser("IE");
			Delay.Milliseconds(100);
			
			
		}
	}
}
