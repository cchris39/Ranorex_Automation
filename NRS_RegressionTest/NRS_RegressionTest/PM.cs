/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-11-18
 * Time: 10:00 AM
 * Scenario: Property Manger Log in and Search Loan file; Update Information: Occupancy, Primary and Alternative phone, Lock Box Code, Status --> 'Sold Firm'.
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
	/// Description of PM.
	/// </summary>
	[TestModule("24B908F1-0735-4B15-9422-B850D066FED8", ModuleType.UserCode, 1)]
	public class PM : ITestModule
	{
		
		public static NRS_RegressionTestRepository repo = NRS_RegressionTestRepository.Instance;
		static PM instance = new PM();
				
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public PM()
		{
			// Do not delete - a parameterless constructor is required!
		}

		#region Varilables
		
		string _varLoan = "";
		[TestVariable("71f5a75c-f3f6-4ff1-9c48-0edb427fc755")]
		public string varLoan
		{
			get { return _varLoan; }
			set { _varLoan = value; }
		}
		
		string _varUrl = "";
		[TestVariable("e8dc5c50-fc09-4ee1-8406-07d9e628e9f5")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varUid = "";
		[TestVariable("f4017abe-a750-4899-bc39-08676142a105")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("211c5ade-e4f5-43eb-a408-5e333a23a6fc")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		#endregion
		
		private string priPhone = "905-0654321";
		private string altPhone = "905-1234560";
		private string lockCode = "2007";
		private string updateStatus = "SOLD_FIRM";
		private string occu = "Tenanted";
		
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		
		/// <summary>
		/// Porperty Manager Update Information
		/// </summary>
		public void pmUpdate(string occu, string phone1, string phone2, string lckCode, string nStatus, string loan)
		{
			//Go to PM information page
			repo.NRS.TopMenu.PropertyManager.MoveTo();
			Delay.Milliseconds(100);
			repo.NRS.PM.PM_PropertyInfo.Click();
			Delay.Milliseconds(500);
						
			//Update Occupancy
			repo.NRS.PM.LoanDataOccupancySelect.TagValue = occu;
			Delay.Milliseconds(100);
			
			//Update Phone Info
			repo.NRS.PM.OwnerPrimaryPhone.PressKeys(phone1);
			repo.NRS.PM.OwnerAlternativePhone.PressKeys(phone2);
			Delay.Milliseconds(100);
			
			//Update Lock Box Code
			repo.NRS.PM.PMLockBoxCode.PressKeys(lckCode);
			Delay.Milliseconds(200);
			
			//Update File Status
			repo.NRS.PM.PmStatus.TagValue = nStatus;
			Delay.Milliseconds(100);
			
			//Update
			repo.NRS.PM.PmStatus_Update.Click();
			Delay.Milliseconds(300);
			
			//Report Update Status
			Validate.Exists(repo.NRS.Record_SuccessfullySavedBox);
			Report.Log(ReportLevel.Success, "Success", "Property Manager information successfully updated for: " + loan);
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
			
			//PM Login; Search file and open file
			Login pm_login = new Login();
			pm_login.usrLogin(varUid, varPwd);
			pm_login.searchFile(varLoan);
			Delay.Milliseconds(100);
			pm_login.openFile();
			
			//Update PM information
			pmUpdate(occu, priPhone, altPhone, lockCode, updateStatus, varLoan);
			Delay.Milliseconds(100);
			
			//Close Browser
			Delay.Milliseconds(100);
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
		}
	}
}
