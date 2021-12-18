/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-11-18
 * Time: 10:46 AM
 * Scenario: Lawyer login and Open file; Update Legal file information and status
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
	/// Description of Lawyer.
	/// </summary>
	[TestModule("86530FAE-96B2-49B7-9D87-C22AF75890D9", ModuleType.UserCode, 1)]
	public class Lawyer : ITestModule
	{
		
		public static NRS_RegressionTestRepository repo = NRS_RegressionTestRepository.Instance;
		static Lawyer instance = new Lawyer();
				
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public Lawyer()
		{
			// Do not delete - a parameterless constructor is required!
		}

		#region Varilables
		
		string _varLoan = "";
		[TestVariable("d58a9662-7944-488a-b94a-137d7390a553")]
		public string varLoan
		{
			get { return _varLoan; }
			set { _varLoan = value; }
		}
		
		string _varUrl = "";
		[TestVariable("919c11a3-d22a-4709-a3bf-54258041fdad")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varUid = "";
		[TestVariable("5e64b664-1ad9-4801-95dd-6632272a01b2")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("d79cd589-385f-4045-8d12-99f08a6e07dc")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
				
		#endregion
		
		private string legalStatus = "DEMAND_ISSUED";
		private string ctNbr;
		private string ordType = "REDEMPTION_ORDER";
		
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		
		/// <summary>
		/// Lawyer update fiel information
		/// </summary>
		public void lawyerUpdate(string courtNbr, string ord, string nStatus, string loan)
		{
			//Go to Legal Information
			repo.NRS.TopMenu.Legal.MoveTo();
			Delay.Milliseconds(100);
			repo.NRS.Lawyer.Lawyer_LegalInfo.Click();
			Delay.Milliseconds(500);
			
			//Update Lawyer information
			repo.NRS.Lawyer.Lawyer_LegalCourtNumber.PressKeys(courtNbr);
			Delay.Milliseconds(200);
			
			repo.NRS.Lawyer.OrderType.TagValue = ord;
			Delay.Milliseconds(200);
			
			//Update status --> 'Demand Issued'
			repo.NRS.Lawyer.LegalStatus.TagValue = nStatus;
			Delay.Milliseconds(100);
			repo.NRS.Lawyer.LawyerStatus_Update.Click();
			Delay.Milliseconds(300);
			
			//Report Status
			Validate.Exists(repo.NRS.Lawyer.SuccessfullySavedRecordLegalInfo);
			Report.Log(ReportLevel.Success, "Success", "Legal Information successfully updated and saved for: " + loan);
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
						
			//Lawyer login; Search and open file
			Login lawyer_login = new Login();
			
			lawyer_login.usrLogin(varUid, varPwd);
			lawyer_login.searchFile(varLoan);
			lawyer_login.openFile();
			Delay.Milliseconds(100);
			
			var random = new Random();
			ctNbr = random.Next(12345, 67890).ToString();
			
			//Lawyer update Legal Information
			lawyerUpdate(ctNbr, ordType, legalStatus, varLoan);
			Delay.Milliseconds(100);
			
			//Close Browser
			Delay.Milliseconds(100);
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
			
		}
	}
}
