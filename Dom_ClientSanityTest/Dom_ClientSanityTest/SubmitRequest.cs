/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 21/12/2015; 04/01/2016
 * Time: 2:09 PM
 * Version 1.0; 1.1 | 
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
	/// Description of SubmitRequest.
	/// Scenario: Client user log into Nas Domestic portal and create a request, report the request status and Nas number.
	/// </summary>
	[TestModule("1BEC6E32-D4C0-4D66-B265-62DFAD09DEF6", ModuleType.UserCode, 1)]
	public class SubmitRequest : ITestModule
	{
		/// <summary>
		/// Using the Dom_SanityTestRepository repository.
		/// </summary>
		public static Dom_SanityTestRepository repo = Dom_SanityTestRepository.Instance;
		
		static SubmitRequest instance = new SubmitRequest();
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public SubmitRequest()
		{
			// Do not delete - a parameterless constructor is required!
		}
		public void KillBrowser(string IE)
		{
		}

		public void ClearBrowserCookies(string IE)
		{
		}
		
		#region Variables
		string _varUid = "";
			[TestVariable("34E146A6-5C1D-4F23-8152-6621B3E94974")]
			public string varUid
			{
				get { return _varUid; }
				set { _varUid = value; }
			}
			
			string _varPwd = "";
			[TestVariable("801F2F27-1FCE-4A3F-BB61-C73D8740909E")]
			public string varPwd
			{
			get { return _varPwd; }
			set { _varPwd = value; }
			}
			
			string _varNasNbr = "";
			[TestVariable("A033FE0C-FD87-4FD9-8514-2041C9CAE304")]
			public string varNasNbr
			{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
			}
					
			string _varRefNbr = "";
			[TestVariable("B7C2E4AD-74C0-4301-9487-7B7E63D7C0E9")]
			public string varRefNbr
			{
			get { return _varRefNbr; }
			set { _varRefNbr = value; }
			}
			
			string _varPropertyAddress = "";
			[TestVariable("5AA6AFCE-2411-42A7-BE87-1D36C71BFD33")]
			public string varPropertyAddress
			{
			get { return _varPropertyAddress; }
			set { _varPropertyAddress = value; }
			}
						
			string _varPostalCode = "";
			[TestVariable("ADC7FA4F-E59F-4A87-9B61-9B526C181BED")]
			public string varPostalCode
			{
				get { return _varPostalCode; }
				set { _varPostalCode = value; }
			}
			
			string _varPurchasePrice = "";
			[TestVariable("6D92D76F-3BA8-4E1F-9DAD-152F597B8871")]
			public string varPurchasePrice
			{
				get { return _varPurchasePrice; }
				set { _varPurchasePrice = value; }
			}
					
			string _varMortgage = "";
			[TestVariable("1CCE6225-92FA-4AFB-9BD6-E3764B3CB47C")]
			public string varMortgage
			{
				get { return _varMortgage; }
				set { _varMortgage = value; }
			}
						
			string _postCode = "";
			[TestVariable("92DDDEAB-1AEC-458E-8924-A5BABF5BC78D")]
			public string postCode
			{
				get { return _postCode; }
				set { _postCode = value; }
			}
			
			
			string _varUrl = "";
			[TestVariable("9D0D8A97-289B-46B3-8C97-D7D8F7E4C83F")]
			public string varUrl
			{
				get { return _varUrl; }
				set { _varUrl = value; }
			}
			
			string _varSerType = "";
			[TestVariable("722EEFDD-764D-46BE-BD8B-DDFC23DD7CEE")]
			public string varSerType
			{
				get { return _varSerType; }
				set { _varSerType = value; }
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
			
			Host.Local.ClearBrowserCookies("IE");
			Delay.Milliseconds(100);
			
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(100);
			
			//Login
			repo.DomNasHome.Userid.PressKeys(varUid);
			repo.DomNasHome.Password.PressKeys(varPwd);
			repo.DomNasHome.RememberUserName.Click();
			repo.DomNasHome.Submit.Click();
			Delay.Milliseconds(200);
			
			if (repo.NAS.NotificationInfo.Exists())
			{repo.NAS.ButtonYes.Click();}
				
				
			//Generate Random Number
			var random = new Random();
			var StrNbr = random.Next(1, 9000);
			var ClientRef = random.Next(123456789, 987654321);
			var Price = random.Next(500,900)*1000;
			var MortgAmount= random.Next(100,500)*1000;
			var ContactNbr1 = random.Next(111111, 999999);
			var ContactNbr2 = random.Next(1111, 9999);
			
			//Convert Integar nunmber to string
			//String StrNbr = Convert.ToString(varStrNbr);
			
			const string varAddressName = "Ranorex Automation Insanity";   //"Ranorex Insanity";
			const string varAddressType = "Test";
				
			string streetNum = Convert.ToString(StrNbr);
			string AppRefNbr = Convert.ToString(ClientRef);	
			string AppPrice = Convert.ToString(Price);	
			string AppMtAmount = Convert.ToString(MortgAmount);	
			string ContactNbr = Convert.ToString(ContactNbr1);	
			string Tel1 = ContactNbr.Substring(0,3);
			string Tel2 = ContactNbr.Substring(3,3);
			string Tel3 = Convert.ToString(ContactNbr2);
			
			varPropertyAddress = streetNum + " " + varAddressName + " " + varAddressType;
			varPostalCode = postCode;
			
			//Order Request
			repo.DomNasHome.RequestService.Click();
			
			repo.DomNasHome.RequestTable.AppPostalCode.PressKeys(varPostalCode);
			repo.DomNasHome.RequestTable.AppPostalCode.PressKeys("{Tab}");
			Delay.Milliseconds(100);
			
			repo.DomNasHome.RequestTable.AppStreetNum.PressKeys(streetNum);
			repo.DomNasHome.RequestTable.AppStreetName.PressKeys(varAddressName);
			repo.DomNasHome.RequestTable.AppStreetType.PressKeys(varAddressType);
			Delay.Milliseconds(100);
			
			//Select Request Service Type
			Ranorex.InputTag reqSerOption = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[@id='service_type' and @value='$']".Replace("$", varSerType).Trim();
			Delay.Milliseconds(100);
			reqSerOption.Click();
			Delay.Milliseconds(100);
			
			repo.DomNasHome.MenuDisplay.Purchase.Click();
			repo.DomNasHome.MenuDisplay.Price.PressKeys(AppPrice);
			repo.DomNasHome.MenuDisplay.MortgageAmt.PressKeys(AppMtAmount);                                          //Not every client use this line
			
			varPurchasePrice = AppPrice;
			varMortgage = AppMtAmount;
			
			repo.DomNasHome.MenuDisplay.AppRefNbr.PressKeys(AppRefNbr);
			repo.DomNasHome.MenuDisplay.AppBraNbr.PressKeys("36888");
			repo.DomNasHome.MenuDisplay.AppName.PressKeys("Teck Smith");
			
			
			if (varSerType != "Drive-By" && varSerType !="Desktop" && varSerType !="Property Gauge" && varSerType!="Market Rent")
			{
			repo.DomNasHome.MenuDisplay.AppContactName.PressKeys("John Smith");
			repo.DomNasHome.MenuDisplay.Tel1.PressKeys(Tel1);
			repo.DomNasHome.MenuDisplay.Tel2.PressKeys(Tel2);
			repo.DomNasHome.MenuDisplay.Tel3.PressKeys(Tel3);
			}
			
			repo.DomNasHome.MenuDisplay.SpDirection.PressKeys("This is a test, please do not proceed; Thanks");
			repo.DomNasHome.MenuDisplay.Invoice.Click();
			repo.DomNasHome.MenuDisplay.SubmitBtn.Click();
			Delay.Milliseconds(100);
			
			repo.DomNasHome.MenuDisplay.SelectAppraiserSubmitBtn.Click();
			Delay.Milliseconds(100);
			
			//Get Nas Number
			var NasNbr = repo.DomNasHome.MenuDisplay.StrongTagViewAppraisalNum.InnerText.Trim();
			NasNbr = NasNbr.Substring(30, 7);
			varNasNbr = NasNbr;
			varRefNbr = AppRefNbr;
			
			//Report request submit status
			Report.Log(ReportLevel.Success, "Validation", "Request has been successful submit");
			Report.Log(ReportLevel.Info, "Validation", "Request Number is: " + NasNbr);
			Validate.Exists(repo.DomNasHome.MenuDisplay.StrongTagYourAppraisalHasBeenSen);
			Delay.Milliseconds(100);
			
			//Close Browser
			//Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
			

		}
		
	}
}
