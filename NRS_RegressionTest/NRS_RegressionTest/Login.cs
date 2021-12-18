/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-11-14
 * Scenario: User log into portal and Log out, verify log in/out successfull
 * Time: 11:24 AM
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

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace NRS_RegressionTest
{
	/// <summary>
	/// Description of Login.
	/// </summary>
	[TestModule("666F0281-B0A3-4CBB-B910-DB096786E544", ModuleType.UserCode, 1)]
	public class Login : ITestModule
	{
		
		public static NRS_RegressionTestRepository repo = NRS_RegressionTestRepository.Instance;
		static Login instance = new Login();
				
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public Login()
		{
			// Do not delete - a parameterless constructor is required!
		}

		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		/// 
		
		#region Varilables
		
		string _varUrl = "";
		[TestVariable("b18d2b5f-0db2-4892-8f33-4ee3c457f5e4")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varUid = "";
		[TestVariable("ce118fff-ac31-4f4b-a0ad-37999ce978c4")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("c3e8dc3b-b67d-4426-b789-8b3a2ef75a20")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		#endregion
		
		public void KillBrowser(string IE)
		{
		}
		
		public void openUrl(string url)
		{
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// User Log in and Report Log in status
		/// </summary>
		public void usrLogin(string uid, string pwd)
		{
			repo.NRS.Homepage.Loginlink.Click();
			Delay.Milliseconds(200);
			
			repo.NRS.Homepage.LoginUsername.PressKeys(uid);
			repo.NRS.Homepage.LoginPassword.PressKeys(pwd);
			repo.NRS.Homepage.LoginBtn.Click();
			
			Validate.Exists(repo.NRS.Logoutlink);
			Report.Log(ReportLevel.Success, "Success", "\"" + uid + "\"" + " login successful");
		}
		
		
		/// <summary>
		/// User Log out and Report Log out status
		/// </summary>
		public void usrLogout()
		{
			repo.NRS.Logoutlink.Click();
			Delay.Milliseconds(100);
			
			Validate.Exists(repo.NRS.Homepage.Loginlink);
			Report.Log(ReportLevel.Success, "Success", "Logout successful");
		}
		
		/// <summary>
		/// Search File
		/// </summary>
		public void searchFile(string loanNbr)
		{
			repo.NRS.Search.Click();
			Delay.Milliseconds(200);
			
			//Input Loan Number and Search
			repo.NRS.QuickSearchField.PressKeys(loanNbr);
			repo.NRS.QuickSaerchIcon.Click();
			Delay.Milliseconds(200);
		}
				
		/// <summary>
		/// Go to File
		/// </summary>
		public void openFile()
		{
			repo.NRS.QuickSearchResult.Click();
			Delay.Milliseconds(400);
		}
			
		
		/// <summary>
		/// Main
		/// </summary>
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(200);

			//User log in and log out
			usrLogin(varUid, varPwd);
			
			Delay.Milliseconds(100);
			
			usrLogout();
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
		
		}
	}
}
