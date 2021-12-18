/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-11-16
 * Time: 11:28 AM
 * Scenario: (1) Verify 'Import/Create Files' tab  (2) Admin Edit Loan data   (3) Reassign Loan to lender   (4) Add Borrower and Update Milestone in 'Master Loan Info'  
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
	/// Description of Admin.
	/// </summary>
	[TestModule("AC20A859-2634-47A1-A686-1D0AE0099A11", ModuleType.UserCode, 1)]
	public class Admin : ITestModule
	{
		public static NRS_RegressionTestRepository repo = NRS_RegressionTestRepository.Instance;
		static Admin instance = new Admin();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public Admin()
		{
			// Do not delete - a parameterless constructor is required!
		}

		#region Varilables
		
		string _varLoan = "";
		[TestVariable("6bf4a22e-aa48-49e5-a78a-49bce58b14d4")]
		public string varLoan
		{
			get { return _varLoan; }
			set { _varLoan = value; }
		}
		
		string _varUid = "";
		[TestVariable("9a0a3d28-a159-4e12-8d41-02e9d523f1f3")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("8d4f5763-5d98-4710-af85-6ae076593934")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varUrl = "";
		[TestVariable("8bff005b-0117-4082-8c90-511be08e0fed")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		#endregion
		
		private string newAmount = "350000.0";
		private string newIntrRate = "0.75";
		private string newLenderValue = "0";
		private string nMilestoneStatus = "2";                            //Milestone: 'Legal';
		private string nBorrowerLstName = "Dump";
		private string nBorrowerFstName = "Donald";
		private string nBorrowerEmail = "airforce1@usb.com";
		
		//private string filePath = @"c:\Upload_Files\toUpload.pdf\";
		
		
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		
		
		/// <summary>
		/// Edit Loan Data
		/// </summary>
		public void editLoanData(string newAmt, string newRate, string loan)
		{
			//Open Loan data to edit
			repo.NRS.TopMenu.AdminlinkSpan.MoveTo();
			Delay.Milliseconds(100);
			
			repo.NRS.Admin.AdminLoanDataTab.MoveTo();
			Delay.Milliseconds(100);
			
			repo.NRS.Admin.AdminEditLoanData.Click();
			Delay.Milliseconds(400);
			
			//Edit 'Original Loan Amount' and 'Interest Rate' field
			repo.NRS.Admin.AdminOriginalLoanAmount.TagValue = newAmt;
			Delay.Milliseconds(100);
			
			repo.NRS.Admin.AdminInterestRate.TagValue = newRate;
			Delay.Milliseconds(200);
			
			repo.NRS.Admin.LoanData_SaveBtn.Click();
			Delay.Milliseconds(400);
			
			//Report Data edit and Saved Status
			Validate.Exists(repo.NRS.Record_SuccessfullySavedBox);
			Report.Log(ReportLevel.Success, "Success", "Loan: '" + loan + "' data updated and saved successfully.");
			Report.Log(ReportLevel.Info, "Information", "New original loan amount: " + newAmt);
			Report.Log(ReportLevel.Info, "Information",  "New interest rate: " + newRate);	                      
		}
		
		/// <summary>
		/// Reassign Loan to Lender
		/// </summary>
		public void reassignLoan(string newVal, string loan)
		{
			//Reassign loan
			repo.NRS.TopMenu.AdminlinkSpan.MoveTo();
			Delay.Milliseconds(100);
			
			repo.NRS.Admin.AdminLoanDataTab.MoveTo();
			Delay.Milliseconds(200);
			
			repo.NRS.Admin.Admin_ReassignLink.Click();
			Delay.Milliseconds(100);
			
			repo.NRS.Admin.ReassignLenderARS.MoveTo();
			repo.NRS.Admin.ReassignLenderARS.PerformClick();
			Delay.Milliseconds(400);
			
			//Select New Lender
			repo.NRS.Admin.AdminSelectLender.TagValue = newVal;
			Delay.Milliseconds(100);
			
			//Reassign Submit
			repo.NRS.Admin.ReassignLenderARS_Btn.Click();
			Delay.Milliseconds(600);
			
			Validate.Exists(repo.NRS.Record_SuccessfullySavedBox);
			Report.Log(ReportLevel.Success, "Success", "Loan: '" + loan + "' successfully reassign to new lender.");
		}
		
		/// <summary>
		/// Master Loan Information section: #Add borrower; #Update Milestone; 
		/// </summary>
		public void addBorrower(string lastName, string firstName, string email)
		{
			//Go to Loan Data 'Master Loan Info' page
			repo.NRS.TopMenu.LoanData.MoveTo();
			Delay.Milliseconds(100);
			
			repo.NRS.LoanData.MasterLoanInfo.Click();
			Delay.Milliseconds(400);
			
			//Add New Borrower
			repo.NRS.LoanData.AddBorrower.Click();
			Delay.Milliseconds(400);
			
			repo.NRS.LoanData.NewBorrower_LastName.PressKeys(lastName);
			Delay.Milliseconds(100);
			repo.NRS.LoanData.NewBorrower_FstName.PressKeys(firstName);
			Delay.Milliseconds(100);
			repo.NRS.LoanData.NewBorrower_Email.PressKeys(email);
			Delay.Milliseconds(100);
			
			repo.NRS.LoanData.NewBorrowerSave.Click();
			Delay.Milliseconds(500);
			
			//Refresh Web Page
			Keyboard.Press("{LControlKey down}{F5 down}{LControlKey up}{F5 up}");
			Delay.Milliseconds(500);
			
			//Validate New Borrower added
			string nBorrower = lastName + "," + firstName;
			string nBorrowerAdd = repo.NRS.LoanData.NewBorrower_Added.InnerText.Trim();
			
			//Report Add New Borrower Status
			Validate.AreEqual(nBorrower, nBorrowerAdd);
			Report.Log(ReportLevel.Success, "Success", "New Borrower added successfully: " + nBorrowerAdd);
		}
		
		
		/// <summary>
		/// Update Milestone status
		/// </summary>
		public void updateMilestone(string newStatus, string loan)
		{
			repo.NRS.LoanData.Milestones.TagValue = newStatus;
			Delay.Milliseconds(200);
			
			//Report Status
			Report.Log(ReportLevel.Info, "Information", "Loan: '" + loan + "' milestone status has been updated to 'Legal'.");
		}
		
		
		/// <summary>
		/// Admin Create Files - Upload
		/// </summary>
		public void verifyImportFileTab()
		{
			//Import/Create Files
			repo.NRS.TopMenu.AdminlinkSpan.MoveTo();
			Delay.Milliseconds(100);
			
			repo.NRS.Admin.ImportCreateFilesTab.MoveTo();
			Delay.Milliseconds(200);
			
			repo.NRS.Admin.ImportCreateFiles.Click();
			Delay.Milliseconds(400);
			
			//Report 'Import/Create Files' tab availability
			Validate.Exists(repo.NRS.Admin.ImportCreateFilesHeader);
			Report.Log(ReportLevel.Info, "Information", "'Import/Create Files' tab verified.");
			Delay.Milliseconds(100);
			
		}
		
		
		
		//Main
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			Delay.Milliseconds(100);
			
			Login admin_login = new Login();
			
			//Verify 'Import/Create files' tab
			verifyImportFileTab();
			Delay.Milliseconds(200);
			
			//Admin Edit Loan Data
			editLoanData(newAmount, newIntrRate, varLoan);
			Delay.Milliseconds(200);
									
			//Reassign loan to lender
			reassignLoan(newLenderValue, varLoan);
			Delay.Milliseconds(200);
			
			//Update Milestone status to 'Legal'  (Value = '2')
			admin_login.searchFile(varLoan);
			Delay.Milliseconds(100);
			
			admin_login.openFile();
			Delay.Milliseconds(200);
			updateMilestone (nMilestoneStatus, varLoan);
			Delay.Milliseconds(200);
			
			//Add New Borrower
			addBorrower(nBorrowerLstName, nBorrowerFstName, nBorrowerEmail);
			Delay.Milliseconds(200);
			
		}
	}
}
