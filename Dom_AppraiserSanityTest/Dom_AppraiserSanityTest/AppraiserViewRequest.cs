/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 26/01/2016
 * Time: 9:18 AM
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

namespace Dom_AppraiserSanityTest
{
	/// <summary>
	/// Description of AppraiserViewRequest.
	/// Scenario: Appraiser search for the assigned (cleint submit) request and view/verify various information 
	/// </summary>
	[TestModule("89EF7A85-4E62-40AB-8D2E-683F7ACB9390", ModuleType.UserCode, 1)]
	public class AppraiserViewRequest : ITestModule
	{
		/// <summary>
		/// Using the Dom_AppraiserSanityTestRepository repository.
		/// </summary>
		public static Dom_AppraiserSanityTestRepository repo = Dom_AppraiserSanityTestRepository.Instance;
		
		static AppraiserViewRequest instance = new AppraiserViewRequest();
		/// <summary>
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public AppraiserViewRequest()
		{
			// Do not delete - a parameterless constructor is required!
		}

		#region Varilables
		
		string _varAppraiser = "";
		[TestVariable("C8DC5221-1F81-4B2B-A549-50F2079684B5")]
		public string varAppraiser
		{
			get { return _varAppraiser; }
			set { _varAppraiser = value; }
		}
		
		string _varPwd = "";
		[TestVariable("9D771AA0-59F9-414D-8A90-FFDA2AF49CA3")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("7B951B9A-8EB4-44CE-B5EB-16C069976EED")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varUrl = "";
		[TestVariable("1AA946DE-1A4A-45FE-BB55-5164B720D7C1")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varPostalCode = "";
		[TestVariable("B7A20901-D634-44F6-B001-C1E7667BDEF9")]
		public string varPostalCode
		{
			get { return _varPostalCode; }
			set { _varPostalCode = value; }
		}
		
		string _varRefNbr = "";
		[TestVariable("BA766DD2-68B0-4BF8-8811-0E69F190C702")]
		public string varRefNbr
		{
			get { return _varRefNbr; }
			set { _varRefNbr = value; }
		}
		
		string _varPropertyAddress = "";
		[TestVariable("27253FCB-3395-4901-82DB-0F3B495C03C3")]
		public string varPropertyAddress
		{
			get { return _varPropertyAddress; }
			set { _varPropertyAddress = value; }
		}
		string _varPurchasePrice = "";
		[TestVariable("B796691F-1B6F-4C46-9797-21439F67E5D4")]
		public string varPurchasePrice
		{
			get { return _varPurchasePrice; }
			set { _varPurchasePrice = value; }
		}
		string _varMortgage = "";
		[TestVariable("39C057B8-4FEC-437F-BDAE-87591B037FDD")]
		public string varMortgage
		{
			get { return _varMortgage; }
			set { _varMortgage = value; }
		}
		
		string _varCdate = "";
		[TestVariable("64CB8DFB-F300-4424-9977-EC11FCBC0D37")]
		public string varCdate
		{
			get { return _varCdate; }
			set { _varCdate = value; }
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
						
			//Host.Local.ClearBrowserCookies("IE");
			//Delay.Milliseconds(100);
			
			//Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			//Delay.Milliseconds(100);
			
			
			//Appraiser log in
			//repo.DomNasHome.Userid.PressKeys(varAppraiser);
			//repo.DomNasHome.Password.PressKeys(varPwd);
			//repo.DomNasHome.Submit.Click();
			Delay.Milliseconds(100);
			
			//Search By Nas Number
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(varNasNbr);     // varNasNbr
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			repo.DomNasHome.MenuDisplay.ViewRequestBtn.Click();
			Delay.Milliseconds(100);
						
			//Get various information from the view request page
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
			//Delay.Milliseconds(200);
		}
	}
}
