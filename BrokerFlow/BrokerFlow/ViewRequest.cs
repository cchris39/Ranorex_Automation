/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 14/03/2016
 * Time: 4:27 PM
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

namespace BrokerFlow
{
	/// <summary>
	/// Description of ViewRequest.
	/// </summary>
	[TestModule("4885DAAB-948E-473A-A69C-FB5AE921F7FD", ModuleType.UserCode, 1)]
	public class ViewRequest : ITestModule
	{
		
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("0B2F2653-F546-4D10-8E95-78D4E6F25223")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varRefNbr = "";
		[TestVariable("74AD2481-C14D-4426-A69E-B54F70A058E8")]
		public string varRefNbr
		{
			get { return _varRefNbr; }
			set { _varRefNbr = value; }
		}
		
		string _varAddress = "";
		[TestVariable("E8441734-395A-4257-A932-4D21B384D04F")]
		public string varAddress
		{
			get { return _varAddress; }
			set { _varAddress = value; }
		}
		
		string _varCdate = "";
		[TestVariable("0D9CE853-4A3F-4052-A39A-667D07A88FA6")]
		public string varCdate
		{
			get { return _varCdate; }
			set { _varCdate = value; }
		}
		
		string _varPostalCode = "";
		[TestVariable("5F16F069-5186-47C5-976C-D684FC8662F0")]
		public string varPostalCode
		{
			get { return _varPostalCode; }
			set { _varPostalCode = value; }
		}
		
		string _varPurchasePrice = "";
		[TestVariable("CF3EDB1B-8DEA-44A3-B72A-0CEC5D39CF46")]
		public string varPurchasePrice
		{
			get { return _varPurchasePrice; }
			set { _varPurchasePrice = value; }
		}
		
		string _varMortgage = "";
		[TestVariable("24F7BF86-B8BF-409F-B25D-BDC5B7EF0979")]
		public string varMortgage
		{
			get { return _varMortgage; }
			set { _varMortgage = value; }
		}
		
		#endregion
		
		
		/// <summary>
		/// Using the Dom_SanityTestRepository repository.
		/// </summary>
		public static BrokerFlowRepository repo = BrokerFlowRepository.Instance;
		
		static ViewRequest instance = new ViewRequest();
				
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ViewRequest()
		{
			// Do not delete - a parameterless constructor is required!
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
			
			//Search By Nas Number
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(varNasNbr);    //varNasNbr
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
				
			//Report Search by Nas Number status
			Report.Log(ReportLevel.Info, "Validation", "Request Number: " + varNasNbr + " was found.");
			///Report.Log(ReportLevel.Success, "Validation", "Request Number: " + varNasNbr +  );
			Validate.Exists(repo.DomNasHome.MenuDisplay.StrongTag1RecordSFound);	
			
			repo.DomNasHome.MenuDisplay.ViewRequestBtn.Click();
			Delay.Milliseconds(100);
						
			//Get various information from the view request page
			var serviceType = repo.DomNasHome.MenuDisplay.AppraisalServiceInfo.InnerText.Trim();
			var nasNbr = repo.DomNasHome.MenuDisplay.NASRequestNumber.InnerText.Trim().Substring(22,10).Trim();
			var address = repo.DomNasHome.MenuDisplay.AddressInfo.InnerText.Trim();
			var postalCode = repo.DomNasHome.MenuDisplay.PostalCodeInfo.InnerText.Trim();
			var price = repo.DomNasHome.MenuDisplay.Priceinformation.InnerText.Trim();
			var finRw3 = repo.DomNasHome.MenuDisplay.FinancialInfoTagRow3.InnerText.Trim();
			var status = repo.DomNasHome.MenuDisplay.StatusInfo.InnerText.Trim();
			var refNbr = repo.DomNasHome.MenuDisplay.ClientRefNbrInfo.InnerText.Trim();
			var cDate = repo.DomNasHome.MenuDisplay.DateCreatedInfo.InnerText.Trim().Substring(0,10);
			var cofDate = repo.DomNasHome.MenuDisplay.COF_Deadline.InnerText.Trim();
			
			string priceAmount = "$" + varPurchasePrice + ".0";
			string mortgageAmount = "$" + varMortgage + ".0";
					
			//Report View request status by validating the information
			Report.Log(ReportLevel.Success, "Validation", "Nas request number is match.");
			Validate.AreEqual(varNasNbr, nasNbr);     //varNasNbr
			
			Report.Log(ReportLevel.Success, "Validation", "Postal Code is match.");
			Validate.AreEqual(varPostalCode, postalCode);     // varPostalCode
				
			Report.Log(ReportLevel.Info, "Service Type is: " + serviceType);
			Report.Log(ReportLevel.Info, "Request Status is: " + status);
			//Report.Log(ReportLevel.Info, "COF Deadline is: " + cofDate);
			
			Report.Log(ReportLevel.Success, "Validation", "Client Reference Number is match.");
			Validate.AreEqual(varRefNbr, refNbr);     //varRefNbr
			
			Report.Log(ReportLevel.Success, "Validation", "Request Date is correct.");
			Validate.AreEqual(varCdate, cDate);     //varCdate
				
			Report.Log(ReportLevel.Success, "Validation", "Purchase Price is match.");
			Validate.AreEqual(priceAmount, price);     // Price
			
			
			if (finRw3 == "COF Deadline:")
			{
				var mortgage = repo.DomNasHome.MenuDisplay.MortgageInformation.InnerText.Trim();
				Report.Log(ReportLevel.Success, "Validation", "Mortgage Amount is match.");
				Validate.AreEqual(mortgageAmount, mortgage);     // mortgage
			
			}else
			{
				var mortgage = repo.DomNasHome.MenuDisplay.COF_Deadline.InnerText.Trim();
				Report.Log(ReportLevel.Success, "Validation", "Mortgage Amount is match.");
				Validate.AreEqual(mortgageAmount, mortgage);     // mortgage
			}
						
			Delay.Milliseconds(200);
			
		}
	}
}
