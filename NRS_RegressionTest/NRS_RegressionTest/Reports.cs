/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-11-17
 * Time: 3:02 PM
 * Scenario: Lender Created Report;  Verify View Report Tab; View Report by openning HTML report and closed report.
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
	/// Description of Reports.
	/// </summary>
	[TestModule("ECE7C558-B882-439C-AF83-6DF56B90CEBE", ModuleType.UserCode, 1)]
	public class Reports : ITestModule
	{
		public static NRS_RegressionTestRepository repo = NRS_RegressionTestRepository.Instance;
		static Reports instance = new Reports();
				
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public Reports()
		{
			// Do not delete - a parameterless constructor is required!
		}

		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		
	
		private const string LOAN_TYPE = "Secured";
		private string startDate, loanSrName;
		
		/// <summary>
		/// Get Loan Servicer name through Quick Menu
		/// </summary>
		static string[] getServicerName()
		{
			string srName = repo.NRS.Reports.ActorName.InnerText.Trim();
			string[] Name = new string[2];
				
			Name = srName.Split(',');
			
			return Name;
		}
		
		/// <summary>
		/// Created Active Report
		/// </summary>
		public void createReport(string srName, string type, string stDate)
		{
			
			string toDate = System.DateTime.Today.ToString("yyyy-MM-dd");
			
			//Go to Reports page
			repo.NRS.TopMenu.Reports.MoveTo();
			Delay.Milliseconds(100);
			repo.NRS.Reports.Reports.Click();
			Delay.Milliseconds(300);
			
			//Create Active Report
			repo.NRS.Reports.ActiveFiles.Click();
			Delay.Milliseconds(300);
			
			repo.NRS.Reports.LoanServicerSelectTag.TagValue = srName;
			Delay.Milliseconds(100);
			
			repo.NRS.Reports.LoanSecuredTypeSelect.TagValue = type;
			Delay.Milliseconds(100);
			
			repo.NRS.Reports.StartDateInputDate.TagValue = stDate;
			Delay.Milliseconds(200);
			
			repo.NRS.Reports.CreateReportBtn.Click();
			Delay.Milliseconds(1500);
			
			if(repo.NRS.Reports.SorryNoDataFoundInfo.Exists(300))
			{
				//Report No data cretaed for Report
				Report.Log(ReportLevel.Warn, "Warning", "No data found for the period of '" + stDate + " to " + toDate + "', no report was created.");
			}else
			{
				//Report Status
				Report.Log(ReportLevel.Info, "Information", "Report cretaed by: '" + srName + "'.");
				Validate.Exists(repo.NRS.Reports.ViewAsPdf);
				Report.Log(ReportLevel.Info, "Information", "'View Preport as PDF' tab verified.");
				Validate.Exists(repo.NRS.Reports.ViewAsExcel);
				Report.Log(ReportLevel.Info, "Information", "'View Preport as Excel' tab verified.");
				Validate.Exists(repo.NRS.Reports.ViewAsHTML);
				Report.Log(ReportLevel.Info, "Information", "'View Preport as HTML' tab verified.");
				
				//Open as HTML report and closed
				openClsReport();
			}
		}
		
		
		/// <summary>
		/// Open and Close Active HTML Report
		/// </summary>
		public void openClsReport()
		{
			//Open and View HTML Report
			repo.NRS.Reports.ViewHTML_ReportLink.Click();
			Delay.Milliseconds(500);
			
			Report.Log(ReportLevel.Info, "Information", "HTML Report opened.");
			
			//Close HTML Report
			repo.NRS.Reports.CloseReport.Click();
			Delay.Milliseconds(300);
			
			Report.Log(ReportLevel.Info, "Information", "HTML Report closed.");
			
		}
		
		
		//Main
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			Delay.Milliseconds(100);
			
			
			startDate = System.DateTime.Today.AddDays(-90).ToString("yyyy-MM-dd");
			
			string[] theName = new string[2];
			theName = getServicerName();
			
			loanSrName = theName[0].Trim() + " " + theName[1].Trim();
			
			//Create Active Report
			createReport(loanSrName, LOAN_TYPE, startDate);
			Delay.Milliseconds(200);
			
			//Close Browser
			Delay.Milliseconds(100);
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
		}
		
		
	}
}
