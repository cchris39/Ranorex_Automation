/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-11-14
 * Scenario: 
 * (1) Lender login and report login status; (2) Report Lender dashboard status; 
 * Time: 11:44 AM
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
	/// Description of Lender.
	/// </summary>
	[TestModule("107C47E5-B5A2-43FA-AC0B-BEBF92CB54AF", ModuleType.UserCode, 1)]
	public class Lender : ITestModule
	{
		
		public static NRS_RegressionTestRepository repo = NRS_RegressionTestRepository.Instance;
		static Lender instance = new Lender();
		
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public Lender()
		{
			// Do not delete - a parameterless constructor is required!
		}

		#region Varilables
		
		string _varUrl = "";
		[TestVariable("c9b95183-0fb5-4877-afd4-abbd12b03dc0")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varUid = "";
		[TestVariable("9733cb48-ab42-48cd-a712-6733b72b595b")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("0a880029-1953-4e91-8f19-7405a64955ff")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		#endregion
		
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		/// 
		
		public void KillBrowser(string IE)
		{
		}
		
		public void openUrl(string url)
		{
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(200);
		}
		
		public void lenderDashboard()
		{
			string fileToday = repo.NRS.LenderDashboard.Div_CretaedToday.InnerText.Trim();
			string dueToday = repo.NRS.LenderDashboard.Div_DueToday.InnerText.Trim();
			string overDue = repo.NRS.LenderDashboard.Div_Overdue.InnerText.Trim();
			string upComing = repo.NRS.LenderDashboard.Div_Upcoming.InnerText.Trim();
			
			//Report Tab Status in dashboard
			Report.Log(ReportLevel.Info, "Information", "Task created today is: " + fileToday);
			Report.Log(ReportLevel.Info, "Information", "Task due today is: " + dueToday);
			Report.Log(ReportLevel.Info, "Information", "Task overdue is: " + overDue);
			Report.Log(ReportLevel.Info, "Information", "Task upcoming is: " + upComing);
						
			Validate.Exists(repo.NRS.LenderDashboard.Milestone_FIP);
			Report.Log(ReportLevel.Info, "Information", "'Milestones for Files in Progress' tab presented.");
			
			Validate.Exists(repo.NRS.LenderDashboard.AutoWIP);
			Report.Log(ReportLevel.Info, "Information", "'Auto WIP' tab presented.");
			Validate.Exists(repo.NRS.LenderDashboard.MWIP_Me);
			Report.Log(ReportLevel.Info, "Information", "'MWIP Created by Me' tab presented.");
			Validate.Exists(repo.NRS.LenderDashboard.MWIP_Others);
			Report.Log(ReportLevel.Info, "Information", "'MWIP Created by Others' tab presented.");
			Validate.Exists(repo.NRS.LenderDashboard.WIP_Task);
			Report.Log(ReportLevel.Info, "Information", "'WIP by Type' tab presented.");
			Validate.Exists(repo.NRS.LenderDashboard.WIP_Provider);
			Report.Log(ReportLevel.Info, "Information", "'WIP by Provider' tab presented.");
		}
		
		//Main
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			//Lender log in
			Login lenderLog = new Login();
			
			//Openr Browser and go to URL
			lenderLog.openUrl(varUrl);
			Delay.Milliseconds(400);
			
			lenderLog.usrLogin(varUid, varPwd);
			Delay.Milliseconds(200);
			
			//Report Lender dashboard status
			lenderDashboard();
			Delay.Milliseconds(500);
		}
	}
}
