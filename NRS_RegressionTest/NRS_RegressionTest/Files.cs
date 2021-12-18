/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-11-15
 * Time: 12:38 PM
 * Scenario: (1) Search File in Progress  (2) Open file  (3) Close file (4) Re-open file
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
	/// Description of Files.
	/// </summary>
	[TestModule("5C5E903E-6826-4117-8AD2-D4C7CEB67212", ModuleType.UserCode, 1)]
	public class Files : ITestModule
	{
		
		public static NRS_RegressionTestRepository repo = NRS_RegressionTestRepository.Instance;
		static Files instance = new Files();
				
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public Files()
		{
			// Do not delete - a parameterless constructor is required!
		}

		#region Varilables
		
		string _varLoan = "";
		[TestVariable("9de6ae04-50ff-44a5-90c3-2da697147023")]
		public string varLoan
		{
			get { return _varLoan; }
			set { _varLoan = value; }
		}
		
		#endregion
		
		private const string CLS_CODE = "PAID_OUT_IN_FULL";             //Closure Code
		private const string mileStoneCode = "5";                       //'Sold'
		public string loan;
		
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		
		/// <summary>
		/// Verify File tabs: (1) Loss Mitigation (2)Legal Action
		/// </summary>
		public void verifyFileTab()
		{
			repo.NRS.TopMenu.Files.MoveTo();
			Delay.Milliseconds(100);
			repo.NRS.Files.FilesInProcess.Click();
			Delay.Milliseconds(100);
			
			Validate.Exists(repo.NRS.Files.LossMitigationTab);
			Report.Log(ReportLevel.Info, "Information", "Loss Mitigation/Secure tab verified.");
			Validate.Exists(repo.NRS.Files.LegalActionTab);
			Report.Log(ReportLevel.Info, "Information", "Legal Action/Secure tab verified.");
		}
				
		/// <summary>
		/// Close and re-open file
		/// </summary>
		public void closeReopenFile(string loan)
		{
			//Switch to 'Legal Action' tab
			repo.NRS.Files.LegalActionTab.Click();
			Delay.Milliseconds(400);
					
			//Search and open File
			repo.NRS.Files.LegalActionLoanNbr_field.PressKeys(loan);
			Delay.Milliseconds(200);
			
			Keyboard.Press("{Tab}");
			//repo.NRS.Search.Click();
			Delay.Milliseconds(800);
						
			Ranorex.ATag searchLoan = "/dom[@domain='uattest.wnrs.ca']//a[@id>'inprocesstablelegalAction:com.nas.recovery.domain.loan.LoanData' and @innertext='$']".Replace("$", loan).Trim();
			Delay.Milliseconds(200);
			searchLoan.Click();
			Delay.Milliseconds(600);
			
			//Verify Milestone Progress Bar
			Validate.Exists(repo.NRS.Files.MilestonesProgressBar);
			Report.Log(ReportLevel.Info, "Information", "File milestones perogress bar verified.");
			Delay.Milliseconds(200);
			
			//Compare loan number is match
			string fileLoanNbr = repo.NRS.Files.FilesLoanNbr.InnerText.Trim();
			Validate.AreEqual(loan, fileLoanNbr);
			Report.Log(ReportLevel.Info, "Information", "Loan file '" + loan + "' opened, loan number is matched.");
			Delay.Milliseconds(100);
			
			//Closed file
			repo.NRS.Files.LoanClosureCodeLegalAction.TagValue = CLS_CODE;
			Delay.Milliseconds(200);
			repo.NRS.Files.ClosedFiles.Click();
			Delay.Milliseconds(300);
			
			Validate.Exists(repo.NRS.Files.SuccessfullySavedRecordBox);
			Report.Log(ReportLevel.Success, "Success", "File: " + varLoan + " closed successfully; Closure Code: '" + CLS_CODE + "'.");
			Delay.Milliseconds(100);
			
			//Re-Open File
			repo.NRS.Files.FilesReopenFile.Click();
			Delay.Milliseconds(400);
			Validate.Exists(repo.NRS.Files.SuccessfullySavedRecordBox);
			Report.Log(ReportLevel.Success, "Success", "File: " + varLoan + " reopen successfully.");
			Delay.Milliseconds(100);
			
		}

		
		//Main		
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			Delay.Milliseconds(100);
			verifyFileTab();
			
			Delay.Milliseconds(200);
			closeReopenFile(varLoan);
			
			Delay.Milliseconds(200);
		}
	}
}
