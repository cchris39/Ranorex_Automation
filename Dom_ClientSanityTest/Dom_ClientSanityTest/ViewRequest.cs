/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 23/12/2015; 04/01/2016
 * Time: 4:20 PM
 * Version: 1.0; 1.1 | 
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
	/// Description of ViewRequest.
	/// Scenario: Validate the apprequest information in View request page against the request submit information
	/// </summary>
	[TestModule("3C101154-9816-410D-B791-8D2F7ADAF765", ModuleType.UserCode, 1)]
	public class ViewRequest : ITestModule
	{
		#region Varilables
		string _varUid = "";
		[TestVariable("DB8779D9-CECD-4E65-A8D5-3BDB30722B3C")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
				
		string _varPwd = "";
		[TestVariable("6A96BD1E-FA71-4A86-83B0-10BAFF45ED35")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("72CC9A99-549E-4A5F-9154-072AC0B7EE18")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varRefNbr = "";
		[TestVariable("07F4AFAF-F8FD-4685-A836-9B06587C8686")]
		public string varRefNbr
		{
			get { return _varRefNbr; }
			set { _varRefNbr = value; }
		}
		
		string _varAddress = "";
		[TestVariable("54E1E290-318E-4479-80AE-B62B8F0A8174")]
		public string varAddress
		{
			get { return _varAddress; }
			set { _varAddress = value; }
		}
		
		string _varCdate = "";
		[TestVariable("1E439C65-9455-457F-82A0-8CC3C4E63567")]
		public string varCdate
		{
			get { return _varCdate; }
			set { _varCdate = value; }
		}
		
		
		string _varPostalCode = "";
		[TestVariable("CEC551C0-898B-4671-837E-E09B2F61FAD1")]
		public string varPostalCode
		{
			get { return _varPostalCode; }
			set { _varPostalCode = value; }
		}
		
		string _varPurchasePrice = "";
		[TestVariable("9B0E261A-93B0-4CE1-910C-325284508BF2")]
		public string varPurchasePrice
		{
			get { return _varPurchasePrice; }
			set { _varPurchasePrice = value; }
		}
				
		string _varMortgage = "";
		[TestVariable("43381824-A2B3-426E-99FF-48B4E6047477")]
		public string varMortgage
		{
			get { return _varMortgage; }
			set { _varMortgage = value; }
		}
		
		string _varUrl = "";
		[TestVariable("1208E208-A5A6-4BCD-8282-875E4F6EE4B1")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
	#endregion
				
		/// <summary>
		/// Using the Dom_SanityTestRepository repository.
		/// </summary>
		public static Dom_SanityTestRepository repo = Dom_SanityTestRepository.Instance;
		
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
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(varNasNbr);    //varNasNbr
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
				
			//Report Search by Nas Number status
			Report.Log(ReportLevel.Info, "Validation", "Request Number: " + varNasNbr + " was found.");
			///Report.Log(ReportLevel.Success, "Validation", "Request Number: " + varNasNbr +  );
			Validate.Exists(repo.DomNasHome.MenuDisplay.StrongTag1RecordSFound);	
			
			repo.DomNasHome.MenuDisplay.ViewRequestBtn.Click();
			Delay.Milliseconds(100);
						
			//Get various informttion from the view request page
			var serviceType = repo.DomNasHome.MenuDisplay.AppraisalServiceInfo.InnerText.Trim();
			var nasNbr = repo.DomNasHome.MenuDisplay.NASRequestNumber.InnerText.Trim().Substring(22,10).Trim();
			
			var address = repo.DomNasHome.MenuDisplay.AddressInfo.InnerText.Trim();
			var postalCode = repo.DomNasHome.MenuDisplay.PostalCodeInfo.InnerText.Trim();
			var price = repo.DomNasHome.MenuDisplay.Priceinformation.InnerText.Trim();
			var mortgage = repo.DomNasHome.MenuDisplay.MortgageInformation.InnerText.Trim();
			var status = repo.DomNasHome.MenuDisplay.StatusInfo.InnerText.Trim();
			var refNbr = repo.DomNasHome.MenuDisplay.ClientRefNbrInfo.InnerText.Trim();
			var cDate = repo.DomNasHome.MenuDisplay.DateCreatedInfo.InnerText.Trim().Substring(0,10);
			
			string priceAmount = "$" + varPurchasePrice + ".0";
			string mortgageAmount = "$" + varMortgage + ".0";
					
			//Report View request status by validating the information
			Report.Log(ReportLevel.Success, "Validation", "Nas request number is match.");
			Validate.AreEqual(varNasNbr, nasNbr);     //varNasNbr
			
			Report.Log(ReportLevel.Success, "Validation", "Postal Code is match.");
			Validate.AreEqual(varPostalCode, postalCode);     // varPostalCode
				
			Report.Log(ReportLevel.Info, "Validation", "Service Type is: " + serviceType);
			Report.Log(ReportLevel.Info, "Validation", "Request Status is: " + status);
			
			Report.Log(ReportLevel.Success, "Validation", "Purchase Price is match.");
			Validate.AreEqual(priceAmount, price);     // Price
			
			Report.Log(ReportLevel.Success, "Validation", "Mortgage Amount is match.");
			Validate.AreEqual(mortgageAmount, mortgage);     // mortgage
			
			Report.Log(ReportLevel.Success, "Validation", "Client Reference Number is match.");
			Validate.AreEqual(varRefNbr, refNbr);     //varRefNbr
			
			Report.Log(ReportLevel.Success, "Validation", "Request Date is correct.");
			Validate.AreEqual(varCdate, cDate);     //varCdate
			
			//Close Browser
			//Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
		}
	}
}
