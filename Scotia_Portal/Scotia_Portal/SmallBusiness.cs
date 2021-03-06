/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-05-16
 * Time: 9:34 AM
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

namespace Scotia_Portal
{
	/// <summary>
	/// Description of SmallBusiness.
	/// </summary>
	[TestModule("2E9E792B-78D3-4A02-A5E8-E892E39D2B0A", ModuleType.UserCode, 1)]
	public class SmallBusiness : ITestModule
	{
		
		public static Scotia_PortalRepository repo = Scotia_PortalRepository.Instance;
		
		static SmallBusiness instance = new SmallBusiness();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public SmallBusiness()
		{
			// Do not delete - a parameterless constructor is required!
		}
		
		private const string addiEmail = "adminuattest.nas.com";
		private const string appName = "Collection Hub";
		private const string contact = "Collection Port";
		private const string spText = "Scotia Small Business test.";
		private const string strName = "North Star";
		private const string strType = "Ave"; 
		
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("4E62B0ED-73EE-4686-96F5-F3A9645659D1")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varComService = "";
		[TestVariable("5B580533-33D4-440C-B1DF-DBCAB36950F6")]
		public string varComService
		{
			get { return _varComService; }
			set { _varComService = value; }
		}
		
		string _varUid = "";
		[TestVariable("23E45A55-0BD9-4C40-9831-DCB9FDA61C87")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("6336E808-073A-4525-85E5-12BEEF4C91AC")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varPoCode = "";
		[TestVariable("D336923A-C18D-4E29-B5E5-3680E516D345")]
		public string varPoCode
		{
			get { return _varPoCode; }
			set { _varPoCode = value; }
		}
		
		string _varLoan = "";
		[TestVariable("238FBBBE-8ECB-4987-97F2-6CF841F6AD4E")]
		public string varLoan
		{
			get { return _varLoan; }
			set { _varLoan = value; }
		}
		
		string _varIVP_ServiceType = "";
		[TestVariable("47B9FD4F-DDD6-47F5-AFE5-D5E969BF0311")]
		public string varIVP_ServiceType
		{
			get { return _varIVP_ServiceType; }
			set { _varIVP_ServiceType = value; }
		}
		
		#endregion
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		/// 
		public void commercialService(string service)
		{
			repo.DomScotia.RequestService.SmallBusiness.Click();
			Delay.Milliseconds(100);
			
			Ranorex.InputTag selCommService = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[#'$']".Replace("$", service).Trim();
			selCommService.Click();
			Delay.Milliseconds(100);
		}
			
		public void requestInfo(string poCode, string stNbr, string stName, string stType, 
		                        string loan, string price, string mortgage, string refNbr, 
		                        string appName, string contact, string tel1, string tel2, 
		                        string tel3, string spTxt, string email)
		{
			//Property Address
			repo.DomScotia.RequestService.PropertyAddress.AppPostalCode.PressKeys(poCode);
			repo.DomScotia.RequestService.PropertyAddress.AppPostalCode.PressKeys("{Tab}");
			Delay.Milliseconds(100);
			
			repo.DomScotia.RequestService.PropertyAddress.AppStreetNum.PressKeys(stNbr);
			repo.DomScotia.RequestService.PropertyAddress.AppStreetName.PressKeys(stName);
			repo.DomScotia.RequestService.PropertyAddress.AppStreetType.PressKeys(stType);
			Delay.Milliseconds(100);
			
			//Loan Purpose
			string loanPath = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[#'$']".Replace("$", loan).Trim();
			Ranorex.InputTag loanOption = loanPath;
			loanOption.Click();
			Delay.Milliseconds(100);
			
			repo.DomScotia.RequestService.LoanPurpose.Price.PressKeys(price);
			repo.DomScotia.RequestService.LoanPurpose.MortgageAmt.PressKeys(mortgage);
			
			//Loan and Contact Information
			repo.DomScotia.RequestService.LoanAndContactInformation.AppRefNbr.PressKeys(refNbr);
			repo.DomScotia.RequestService.LoanAndContactInformation.AppName.PressKeys(appName);
						
			repo.DomScotia.RequestService.LoanAndContactInformation.AppContactName.PressKeys(contact);
			repo.DomScotia.RequestService.LoanAndContactInformation.Tel1.PressKeys(tel1);
			repo.DomScotia.RequestService.LoanAndContactInformation.Tel2.PressKeys(tel2);
			repo.DomScotia.RequestService.LoanAndContactInformation.Tel3.PressKeys(tel3);
			repo.DomScotia.RequestService.LoanAndContactInformation.PreferNumber1.Click();
			Delay.Milliseconds(100);
			//End of Loan and Contact Information section
			
			//Email notifications
			repo.DomScotia.RequestService.EmailNotifications.AdditionalEmail.PressKeys(email);
			Delay.Milliseconds(100);
			
			//Select Payment type as "Invoice"   (Added @ 2016.05.16)
			repo.DomScotia.RequestService.Payment.Invoice.Click();
			
			//Submit request
			repo.DomScotia.RequestService.SubmitRequest.MoveTo();
			Delay.Milliseconds(100);
			repo.DomScotia.RequestService.SubmitRequest.PerformClick();
			Delay.Milliseconds(500);
		}
		
		public string getRequestNumber()
		{
			string nbr = repo.DomScotia.StrongTagAppraisalRequestNumber.InnerText.Trim().Substring(30, 7);
			return nbr;
			
		}    //End of [reportOrderStatus] function
		
		public string getRequestServiceType()
		{
			repo.DomScotia.StrongTagAppraisalRequestNumber.Click();
			Delay.Milliseconds(100);
			repo.DomScotia.SearchResults.ViewAppraisalRequest.Click();
			Delay.Milliseconds(200);
			
			string serType = repo.DomScotia.ViewRequest.ServiceType.InnerText.Trim();
			return serType;
			
		}    //End of [reportOrderStatus] function
		
		
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			Delay.Milliseconds(200);
			
			var random = new Random();
			string stNbr = random.Next(1, 900).ToString();
			string clientRef = random.Next(123456789, 987654321).ToString();
			int Price = random.Next(500,900)*1000;
			int MortgAmount= random.Next(100,500)*1000;
			string contactNbr1 = random.Next(111, 999).ToString();
			string contactNbr2 = random.Next(111, 999).ToString();
			string contactNbr3 = random.Next(1111,9999).ToString();
			string price = Price.ToString();
			string mortgage = MortgAmount.ToString();
			
			Login UsrLogin = new Login();
			UsrLogin.lauchScotia();
			UsrLogin.UserLogin(varUid, varPwd);
			
			//Request New Service
			repo.DomScotia.MainMenu.RequestService.Click();
			Delay.Milliseconds(100);
			
			commercialService(varComService);
			
			requestInfo(varPoCode, stNbr, strName, strType, varLoan, price, mortgage, clientRef, appName, contact, contactNbr1, contactNbr2, contactNbr3, spText, addiEmail);
			Delay.Milliseconds(100);
			
			//Get Nas number
			varNasNbr = getRequestNumber();
			string reqServiceType = getRequestServiceType();
						
			Report.Log(ReportLevel.Info, varNasNbr + " request service type is: " + reqServiceType);
			Validate.AreEqual(reqServiceType, varIVP_ServiceType);
			
			Delay.Milliseconds(200);
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
			
			
		}
	}
}
