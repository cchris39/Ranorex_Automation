/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-11-14
 * Time: 12:50 PM
 * Scenario: Lender WIP (Work in Progress) section: Lender choose 1st Loan and verified loan data by viewing
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
	/// Description of WIP.
	/// </summary>
	[TestModule("243BFB99-B058-42AE-B27B-F007251C4E65", ModuleType.UserCode, 1)]
	public class WIP : ITestModule
	{
		public static NRS_RegressionTestRepository repo = NRS_RegressionTestRepository.Instance;
		static Lender instance = new Lender();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public WIP()
		{
			// Do not delete - a parameterless constructor is required!
		}

		
		#region Varilables
		
		string _varLoan = "";
		[TestVariable("c57d4495-604e-47f4-b434-14f3db0cb2f5")]
		public string varLoan
		{
			get { return _varLoan; }
			set { _varLoan = value; }
		}
		
		#endregion
		
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		/// 
		
		private string loanNbr, borrower, address, dateCreated, dateDue, task;
		private string viewLoanNbr, viewBorrower, viewAddress, viewDateCreated, viewDateDue, viewTask;
		
		
		/// <summary>
		/// Compare string array
		/// </summary>
		static bool ArraysAreEqual(string[] x, string[] y)
		    {
            if (x.Length != y.Length) {
                return false;
            }

            for (int index = 0; index < x.Length; index++)
            {
                if (x[index] != y[index]) {
                    return false;
                }
            }

            return true;
        	}
		
		
		public static string PutIntoQuotes(string value)
        {
            return "\"" + value + "\"";
        }
		
		
		/// <summary>
		/// Verifiy WIP Tabs
		/// </summary>
		public void wip_VerifyTab()
		{
			repo.NRS.TopMenu.WIP.Click();
			Delay.Milliseconds(200);
			
			//Verify All WIP tab
			Validate.Exists(repo.NRS.WIP.Lender_WIP);
			Report.Log(ReportLevel.Info, "Information", "Lender's Current File WIP tab verified.");
			Validate.Exists(repo.NRS.WIP.Lender_AllFile_WIP);
			Report.Log(ReportLevel.Info, "Information", "Lender's All File WIP tab verified.");
			Validate.Exists(repo.NRS.WIP.LenderOutstandingWipTab);
			Report.Log(ReportLevel.Info, "Information", "Lender's Outsatnding WIP tab verified.");
		}
			
		
		/// <summary>
		/// Get 1st loan task and view
		/// </summary>
		public void wip_1stLoan()
		{	
			string[] fstLoan = new string[5];
			string[] viewFstLoan = new string[5];
			
			//Select Lender user
			repo.NRS.WIP.UsrSelectionBox.TagValue = "1";
			repo.NRS.WIP.SelectBtn.Click();
			Delay.Milliseconds(100);
			
			//Get 1st loan data
			repo.NRS.WIP.N1stLoan_Nbr.Click();
			Delay.Milliseconds(200);
			fstLoan = loan_wip();
			
			//Past loan number to Global parameter
			varLoan = fstLoan[0];
							
			//View 1st Loan
			//repo.NRS.WIP.N1stLoan_View.Click();
			repo.NRS.WIP.WIP_View.Click();
			Delay.Milliseconds(300);
			
			viewFstLoan = viewLoan();
			
			//Compare Loan information with string array
			bool assert1 = ArraysAreEqual(fstLoan, viewFstLoan);
			if(assert1)
			{
				Report.Log(ReportLevel.Success, "Success", "Loan " + fstLoan[0] + " information verified, information matched.");
				Report.Log(ReportLevel.Info, "Information", "Loan number: " + PutIntoQuotes(fstLoan[0])); 
				Report.Log(ReportLevel.Info, "Information", "Borrower: " + PutIntoQuotes(fstLoan[1])); 
				Report.Log(ReportLevel.Info, "Information", "Property Address: " + PutIntoQuotes(fstLoan[2])); 
				Report.Log(ReportLevel.Info, "Information", "Date Created: " + PutIntoQuotes(fstLoan[3])); 
				//Report.Log(ReportLevel.Info, "Information", "Date Due: " + PutIntoQuotes(fstLoan[4])); 
				Report.Log(ReportLevel.Info, "Information", "Task name: " + PutIntoQuotes(fstLoan[4])); 
			}else
			{ Report.Log(ReportLevel.Failure, "Fail", "Loan: " + fstLoan[0] + " information verified, information not matched."); }
			
		}
		
		/// <summary>
		/// Loan data array
		/// </summary>
		public string [] loan_wip()
		{
			string loanNbr = repo.NRS.WIP.N1stLoan_Nbr.InnerText.Trim();
			string borrower = repo.NRS.WIP.N1stLoan_Borrower.InnerText.Trim();
			string address = repo.NRS.WIP.LoanAddress.InnerText.Trim();
			string dateCreated = repo.NRS.WIP.N1stLoan_Created.InnerText.Trim();
			//string dateDue = repo.NRS.WIP.N1stLoan_Due.InnerText.Trim();
			string task = repo.NRS.WIP.N1stLoan_Task.InnerText.Trim();
			
			string[] fstLoan = new string[] {loanNbr, borrower, address, dateCreated, task};
			return fstLoan;
		}
		
		/// <summary>
		/// View 1st loan data array
		/// </summary>
		public string[] viewLoan()
		{
			string viewLoanNbr = repo.NRS.WIP.ViewLoanNbr.InnerText.Trim();
			string viewBorrower = repo.NRS.WIP.ViewBorrower.InnerText.Trim();
			string viewAddress = repo.NRS.WIP.Task_Address.InnerText.Trim();
			string viewDateCreated = repo.NRS.WIP.TaskDate_Created.InnerText.Trim();
			//string viewDateDue = repo.NRS.WIP.TaskDueDate.InnerText.Trim();
			string viewTask = repo.NRS.WIP.TaskHeader_Task.InnerText.Remove(0,5).Trim();
			
			string[] viewFstLoan = new string[] {viewLoanNbr, viewBorrower, viewAddress, viewDateCreated, viewTask};
			return viewFstLoan;
			
		}
		
		
		
		
		//Main
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			Delay.Milliseconds(100);
			
			//Open WIP tab
			wip_VerifyTab();
			Delay.Milliseconds(300);
			
			wip_1stLoan();
			Delay.Milliseconds(200);
		}
	}
}
