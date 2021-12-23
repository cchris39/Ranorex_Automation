/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-05-05
 * Time: 11:21 AM
 * 
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

namespace Scotia_Portal
{
	/// <summary>
	/// Description of Login.
	/// </summary>
	[TestModule("FE3C2FAF-7944-479C-A5F4-5943A24A70CA", ModuleType.UserCode, 1)]
	public class Login : ITestModule
	{
		
		public static Scotia_PortalRepository repo = Scotia_PortalRepository.Instance;
		
		static Login instance = new Login();
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public Login()
		{
			// Do not delete - a parameterless constructor is required!
		}
		
		#region Varilables
		
		string _varUid = "";
		[TestVariable("D2F04DA3-7974-4D75-B0D7-A7778CC14212")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("E2A950B3-EE9E-4557-8CF9-C6E760E21B3C")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		#endregion
		
		 public void KillBrowser(string IE)
		{
		}

		public void ClearBrowserCookies(string IE)
		{
		}
		public void UserLogin (string uid, string pwd)
		{
			repo.DomScotia.Welcome.Click();
			repo.DomScotia.Username.PressKeys(uid);
			repo.DomScotia.Password.PressKeys(pwd);
			repo.DomScotia.Login.Click();
			Delay.Milliseconds(100);
		}
		
		public void lauchScotia()
		{
			Host.Local.OpenBrowser("http://uattest.nas.com/NAS/scotiabanknas", "IE", "", false, false, false, false, false);
			Delay.Milliseconds(100);
				
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
			
			lauchScotia();
			UserLogin(varUid, varPwd);
			
			//Close Browser
			//Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
		}
	}
}
