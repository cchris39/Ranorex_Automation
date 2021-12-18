/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 21/12/2015
 * Time: 12:17 PM
 * Version 1.0
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

namespace Dom_ClientSanityTest
{
	/// <summary>
	/// Description of Login.
	/// Scenario: Client user log into Nas Dom Portal and Log out; Verifying Log in and log out status 
	/// </summary>
	[TestModule("C4E0DCC4-5BDF-4062-A322-BF2FD623D118", ModuleType.UserCode, 1)]
	public class Login : ITestModule
	{	
		/// <summary>
		/// Using the Dom_SanityTestRepository repository.
		/// </summary>
		public static Dom_SanityTestRepository repo = Dom_SanityTestRepository.Instance;

		static Login instance = new Login();
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public Login()
		{
			// Do not delete - a parameterless constructor is required!
		}
		public void KillBrowser(string IE)
		{
		}

		public void ClearBrowserCookies(string IE)
		{
		}
#region Variables
		string _varUid = "";
		[TestVariable("C0DAAF40-14EF-4537-B578-1D58314E4018")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}

		string _varPwd = "";
		[TestVariable("886E39C8-CFC2-465E-8DCB-E5C002BA4C0C")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		
		string _varUrl = "";
		[TestVariable("5B8180B2-E60A-423F-9945-8BD17EB3584A")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
#endregion

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
				
			Host.Local.ClearBrowserCookies("IE");
			Delay.Milliseconds(100);
			
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(0);
			
			repo.DomNasHome.Userid.PressKeys(varUid);
			repo.DomNasHome.Password.PressKeys(varPwd);
			repo.DomNasHome.Submit.Click();
			
			//Report Log in status
			Report.Log(ReportLevel.Success, "Validation", "Login successfull.");
			Validate.Exists(repo.DomNasHome.LogOutInfo);
			
			
			repo.DomNasHome.LogOut.Click();
			Delay.Milliseconds(100);	
			
			//Report Log in status
			Report.Log(ReportLevel.Success,"Validation","Logout Successfull");
			Validate.Exists(repo.DomNasHome.SubmitInfo);
			Validate.Attribute(repo.DomNasHome.SubmitInfo, "Value", "Login");
			
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(100);
			
		}
	}
}
