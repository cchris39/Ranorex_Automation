/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 07/01/2016
 * Time: 4:39 PM
 * Version 1.0
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

namespace Dom_ClientSanityTest
{
	/// <summary>
	/// Description of ChangeRequestStatus.
	/// Scenario: Client change the request status by cancelling the order.
	/// </summary>
	[TestModule("1109263B-A5CE-4C08-9CC4-648AF4600634", ModuleType.UserCode, 1)]
	public class ChangeRequestStatus : ITestModule
	{
		/// <summary>
		/// Using the Dom_SanityTestRepository repository.
		/// </summary>
		public static Dom_SanityTestRepository repo = Dom_SanityTestRepository.Instance;
		
		static ChangeRequestStatus instance = new ChangeRequestStatus();
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ChangeRequestStatus()
		{
			// Do not delete - a parameterless constructor is required!
		}
			
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public void KillBrowser(string IE)
		{
		}

		public void ClearBrowserCookies(string IE)
		{
		}
#region Variables

		string _varUid = "";
		[TestVariable("814602CD-34EE-44F0-B90D-B410987EA9BF")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("531A05CC-CE4C-4F81-A80C-715144D44054")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("DC36951C-7586-402F-8760-BB8C40E75AD9")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varUrl = "";
		[TestVariable("C47CC097-1522-48FE-B4D0-963FD2BE5AB4")]
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
			
			
			/*/
			Host.Local.ClearBrowserCookies("IE");
			Delay.Milliseconds(100);
			
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(200);
			
			//Login
			repo.DomNasHome.Userid.PressKeys(varUid);
			repo.DomNasHome.Password.PressKeys(varPwd);
			repo.DomNasHome.Submit.Click();
			Delay.Milliseconds(100);
			/*/
			
			//Search By Nas Number
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(varNasNbr);     // varNasNbr
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			//Get current status from search result
			var status = repo.DomNasHome.MenuDisplay.StrongTagStatus.InnerText.Trim();
			const string changeStatus = "Cancelled";
			
			//Change request status 
				if (status != "Completed") 
				{
				repo.DomNasHome.MenuDisplay.CancelledBtn.Click();
				repo.DomNasHome.MenuDisplay.ButtonTagYes.Click();
				Delay.Milliseconds(200);
				
				var chgStatus = repo.DomNasHome.MenuDisplay.StrongTagStatus.InnerText.Trim();
				
				//report change status
				Report.Log(ReportLevel.Success, "Validation", "Request has been successfully cancelled.");
				Report.Log(ReportLevel.Info, "Validation", varNasNbr + "Current status is: " + chgStatus);     //varNasNbr
				Validate.AreEqual(changeStatus, chgStatus);
				Delay.Milliseconds(100);
				}
				else if (status == "Completed")
				{
				Report.Log(ReportLevel.Info, "Warning", "Request status is completed, it can not be cancelled.");
				}
				else 
				{
				Report.Log(ReportLevel.Info, "Validation", varNasNbr + " " + "Current status is: " + status);     //varNasNbr
				Report.Log(ReportLevel.Failure, "Validation", "Request has not been cancelled.");
				Validate.NotExists(repo.DomNasHome.MenuDisplay.StatusChangedFromNewToCancelledFor);
				Delay.Milliseconds(100);
				}
		    
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
		}
			
	}
}
