/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 11/01/2016
 * Time: 11:37 AM
 * Version 1.0
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

namespace Dom_AppraiserSanityTest
{
	/// <summary>
	/// Description of ClientSubmitRequest.
	/// </summary>
	[TestModule("865EC56F-D5F3-4F22-B78A-F26C4F5D5509", ModuleType.UserCode, 1)]
	public class ClientSubmitRequest : ITestModule
	{
		/// <summary>
		/// Using the Dom_AppraiserSanityTestRepository repository.
		/// </summary>
		public static Dom_AppraiserSanityTestRepository repo = Dom_AppraiserSanityTestRepository.Instance;
		
		static ClientSubmitRequest instance = new ClientSubmitRequest();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ClientSubmitRequest()
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
		
		string _varClientID = "";
		[TestVariable("B559C129-ADBD-486C-8F5A-7A407C14561C")]
		public string varClientID
		{
			get { return _varClientID; }
			set { _varClientID = value; }
		}
		
		string _varPwd = "";
		[TestVariable("3F33A48A-2637-49CE-900B-F51851521715")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varUrl = "";
		[TestVariable("C1C39709-A028-46B0-A0A6-7D34C338BDD5")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varPostalCode = "";
		[TestVariable("5D6278AB-F875-493D-B960-22327BA933AF")]
		public string varPostalCode
		{
			get { return _varPostalCode; }
			set { _varPostalCode = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("4AADA8E9-453B-48AE-B67F-77983D5ED4E8")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varRefNbr = "";
		[TestVariable("B9323745-411C-4B77-83CB-423250D5D0D6")]
		public string varRefNbr
		{
			get { return _varRefNbr; }
			set { _varRefNbr = value; }
		}
		
		string _varPropertyAddress = "";
		[TestVariable("2CB4C895-4F2E-4423-9B44-AA25160646A7")]
		public string varPropertyAddress
		{
			get { return _varPropertyAddress; }
			set { _varPropertyAddress = value; }
		}
		
		string _varPurchasePrice = "";
		[TestVariable("E5B003B1-484A-4E8E-97F9-85BBED26524C")]
		public string varPurchasePrice
		{
			get { return _varPurchasePrice; }
			set { _varPurchasePrice = value; }
		}
		
		string _varMortgage = "";
		[TestVariable("5E15DB90-EE57-4252-9C2B-92D92A006C60")]
		public string varMortgage
		{
			get { return _varMortgage; }
			set { _varMortgage = value; }
		}
		
		
		string _varCdate = "";
		[TestVariable("D7EB76B9-7516-4171-B8D0-4496A5714C16")]
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
			
			Host.Local.ClearBrowserCookies("IE");
			Delay.Milliseconds(100);
			
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(100);
			
			//client log in
			repo.DomNasHome.Userid.PressKeys(varClientID);
			repo.DomNasHome.Password.PressKeys(varPwd);
			repo.DomNasHome.Submit.Click();
			Delay.Milliseconds(1000);
			
			if (repo.PasswordNotificationInfo.Exists())
			{repo.SavePwdNo.Press();}
			
			/// <summary>
			//client submit request
			/// <summary>
			//Generate Random Number
			var random = new Random();
			var StrNbr = random.Next(1, 9000);
			var ClientRef = random.Next(123456789, 987654321);
			var Price = random.Next(500,900)*1000;
			var MortgAmount= random.Next(100,500)*1000;
			var ContactNbr1 = random.Next(111111, 999999);
			var ContactNbr2 = random.Next(1111, 9999);
			var toDate = System.DateTime.Today.ToString("yyyy-MM-dd");
			
			varCdate = toDate;
			
			//Convert Integar nunmber to string
			//String StrNbr = Convert.ToString(varStrNbr);
			
			const string varAddressName = "Ranorex Insanity";   
			const string varAddressType = "Court";
				
			string streetNum = Convert.ToString(StrNbr);
			string AppRefNbr = Convert.ToString(ClientRef);	
			string AppPrice = Convert.ToString(Price);	
			string AppMtAmount = Convert.ToString(MortgAmount);	
			string ContactNbr = Convert.ToString(ContactNbr1);	
			string Tel1 = ContactNbr.Substring(0,3);
			string Tel2 = ContactNbr.Substring(3,3);
			string Tel3 = Convert.ToString(ContactNbr2);
			
			varPropertyAddress = streetNum + " " + varAddressName + " " + varAddressType;
			//Order Request
			repo.DomNasHome.RequestService.Click();
			
			repo.DomNasHome.RequestTable.AppPostalCode.PressKeys(varPostalCode);
			repo.DomNasHome.RequestTable.AppPostalCode.PressKeys("{Tab}");
			Delay.Milliseconds(100);
			
			repo.DomNasHome.RequestTable.AppStreetNum.PressKeys(streetNum);
			repo.DomNasHome.RequestTable.AppStreetName.PressKeys(varAddressName);
			repo.DomNasHome.RequestTable.AppStreetType.PressKeys(varAddressType);
			Delay.Milliseconds(100);
			
			repo.DomNasHome.MenuDisplay.ServiceTypeFull.Click();
			Delay.Milliseconds(100);
			
			repo.DomNasHome.MenuDisplay.Purchase.Click();
			repo.DomNasHome.MenuDisplay.Price.PressKeys(AppPrice);
			repo.DomNasHome.MenuDisplay.MortgageAmt.PressKeys(AppMtAmount);
			
			varPurchasePrice = AppPrice;
			varMortgage = AppMtAmount;
			
			repo.DomNasHome.MenuDisplay.AppRefNbr.PressKeys(AppRefNbr);
			repo.DomNasHome.MenuDisplay.AppBraNbr.PressKeys("666888");
			repo.DomNasHome.MenuDisplay.AppName.PressKeys("Dan Smith");
			
			repo.DomNasHome.MenuDisplay.AppContactName.PressKeys("John Smith");
			repo.DomNasHome.MenuDisplay.Tel1.PressKeys(Tel1);
			repo.DomNasHome.MenuDisplay.Tel2.PressKeys(Tel2);
			repo.DomNasHome.MenuDisplay.Tel3.PressKeys(Tel3);
			
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
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
			
		}
	}
}
