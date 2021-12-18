/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-11-18
 * Time: 11:23 AM
 * Scenario: Realtor Log in, search and open file; Update file information and status.
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace NRS_RegressionTest
{
	/// <summary>
	/// Description of Realtor.
	/// </summary>
	[TestModule("F73BC6EF-6C4E-423D-839B-5302DDD4750B", ModuleType.UserCode, 1)]
	public class Realtor : ITestModule
	{
		
		public static NRS_RegressionTestRepository repo = NRS_RegressionTestRepository.Instance;
		static Realtor instance = new Realtor();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public Realtor()
		{
			// Do not delete - a parameterless constructor is required!
		}

		#region Varilables
		
		string _varLoan = "";
		[TestVariable("0d189daf-410f-4cf3-8b3f-25676f104719")]
		public string varLoan
		{
			get { return _varLoan; }
			set { _varLoan = value; }
		}
		
		string _varUrl = "";
		[TestVariable("4247aefe-a8fe-41b7-b833-243defae7bd9")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varUid = "";
		[TestVariable("34d4e627-ebe0-4c25-87fc-825a8190e033")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("02b4b560-e517-47da-b8ad-8f09c27f29c5")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		#endregion
		
		private string boardValue = "37";                      // 37- Toronto
		private string nPrice = "950000.0";
		private string nStatus = "LISTED";
		
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		
		/// <summary>
		/// Realtor update information and file status
		/// </summary>
		public void realtorUpdate(string board, string price, string status, string loan)
		{
			//Go to Real Esate Listing page
			repo.NRS.TopMenu.Realestate.MoveTo();
			Delay.Milliseconds(100);
			
			repo.NRS.Realtor.AgentListing.Click();
			Delay.Milliseconds(600);
			
			if(repo.NRS.Realtor.NoDataInfo.Exists(200))
			{
				//Add Real Esate Listing
				repo.NRS.Realtor.AddListing.Click();
			}else
			{
			
				//Open File to update
				repo.NRS.Realtor.QuickEdit.Click();
				Delay.Milliseconds(200);
			}
				//Update information: Board --> Toronto, Price--> New Price, Status --> Listed
				repo.NRS.Realtor.RealEstateBoard.TagValue = board;
				Delay.Milliseconds(100);
			
				repo.NRS.Realtor.ListPrice.TagValue = price;
				Delay.Milliseconds(100);
				repo.NRS.Realtor.MinSalePrice.TagValue = price;
				Delay.Milliseconds(100);
			
				repo.NRS.Realtor.SaveListingBtn.Click();
				Delay.Milliseconds(300);
			
				//Update Listing status to: 'Listed'
				repo.NRS.Realtor.RealtorStatus.TagValue = status;
				Delay.Milliseconds(100);
				repo.NRS.Realtor.Realtor_Update.Click();
				Delay.Milliseconds(300);
			
				//Report Update Status
				Validate.Exists(repo.NRS.Record_SuccessfullySavedBox);
				Report.Log(ReportLevel.Success, "Success", "Real Estate Listing information update and saved for: " + loan);
			
		}
		
		
		
		//Main
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			Delay.Milliseconds(100);
			
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(300);
			
			//Realtor login adn search/open file
			Login realtor_login = new Login();
			
			realtor_login.usrLogin(varUid, varPwd);
			realtor_login.searchFile(varLoan);
			Delay.Milliseconds(100);
			
			realtor_login.openFile();
			Delay.Milliseconds(100);
			
			//Realtor Update information
			realtorUpdate(boardValue, nPrice, nStatus, varLoan);
			Delay.Milliseconds(200);
			
			//Close Browser
			Delay.Milliseconds(100);
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
			
		}
	}
}
