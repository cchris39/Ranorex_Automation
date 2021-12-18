/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 18/02/2016
 * Time: 9:36 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using System.Data;
using MySql.Data.MySqlClient;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace BrokerFlow
{
	/// <summary>
	/// Description of SubmitRequest.
	/// </summary>
	[TestModule("A0EC1B08-7F43-459B-A2CF-46A16A8723FF", ModuleType.UserCode, 1)]
	public class SubmitRequest : ITestModule
	{
		/// <summary>
		/// Using the Dom_AppraiserSanityTestRepository repository.
		/// </summary>
		public static BrokerFlowRepository repo = BrokerFlowRepository.Instance;
		
		static SubmitRequest instance = new SubmitRequest();
		
		#region Varilables
		
		string _varUrl = "";
		[TestVariable("3A9C8BA6-0D62-4565-807C-8DFD176CB206")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varUid = "";
		[TestVariable("DAB9740F-65D4-49C2-8B87-01BF421325BC")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("B3B83E98-C486-42B9-B338-56017C5E7CB9")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varPostalCode = "";
		[TestVariable("9355285A-4EA7-4D1F-9425-B32DBCE987F2")]
		public string varPostalCode
		{
			get { return _varPostalCode; }
			set { _varPostalCode = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("4AD5B40F-4CA3-442F-B874-4DACB10A28F0")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varLender = "";
		[TestVariable("ACEDC04E-6DDD-4DFE-96A9-2E0B25F57676")]
		public string varLender
		{
			get { return _varLender; }
			set { _varLender = value; }
		}
		
		string _varSerType = "";
		[TestVariable("9DDEBC76-6965-462C-AC35-7F6823B11719")]
		public string varSerType
		{
			get { return _varSerType; }
			set { _varSerType = value; }
		}
		
		string _varPropertyAddress = "";
		[TestVariable("EA8735F4-A28A-4CC0-966B-62FBEDCC4E88")]
		public string varPropertyAddress
		{
			get { return _varPropertyAddress; }
			set { _varPropertyAddress = value; }
		}
		
		
		string _varRefNbr = "";
		[TestVariable("33EBA2FB-8212-4FD9-8611-F8FAA7380F88")]
		public string varRefNbr
		{
			get { return _varRefNbr; }
			set { _varRefNbr = value; }
		}
		
		string _varPrice = "";
		[TestVariable("0DC59D4C-56B3-493B-BC2A-4AF67156538E")]
		public string varPrice
		{
			get { return _varPrice; }
			set { _varPrice = value; }
		}
				
		string _varMortgage = "";
		[TestVariable("377B76AE-F0E8-49D7-9491-8CCE2CC5487D")]
		public string varMortgage
		{
			get { return _varMortgage; }
			set { _varMortgage = value; }
		}
		
		string _varCreationDate = "";
		[TestVariable("9476911E-F303-4CB5-9344-630BD9469C83")]
		public string varCreationDate
		{
			get { return _varCreationDate; }
			set { _varCreationDate = value; }
		}
		
		string _PostalCode = "";
		[TestVariable("E2CF5D4D-823C-4528-8E46-453F5C3ACEF8")]
		public string PostalCode
		{
			get { return _PostalCode; }
			set { _PostalCode = value; }
		}
		
		#endregion
	
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

			//WebDocument DomWebDocu = "/dom[@domain='uattest.nas.com']";

			//Broker log in
			repo.DomNasHome.Userid.PressKeys(varUid);
			repo.DomNasHome.Password.PressKeys(varPwd);
			repo.DomNasHome.RememberUserName.TagValue = "No";
			repo.DomNasHome.RememberUserName.Click();
			repo.DomNasHome.Submit.Click();
			Delay.Milliseconds(300);
			
			if(repo.PasswordNotificationInfo.Exists())
				{
				repo.SavePwdNo.Press();
				Delay.Milliseconds(100);
				}
			
			//Generate Random Number
			var random = new Random();
			var StrNbr = random.Next(1, 9000);
			var ClientRef = random.Next(123456789, 987654321);
			var Price = random.Next(500,900)*1000;
			var MortgAmount= random.Next(100,500)*1000;
			var ContactNbr1 = random.Next(111111, 999999);
			var ContactNbr2 = random.Next(1111, 9999);
			
			const string varAddressName = "Broker Automation";   //"Ranorex Insanity";
			const string varAddressType = "Test";
				
			string streetNum = Convert.ToString(StrNbr);
			string AppRefNbr = Convert.ToString(ClientRef);	
			string AppPrice = Convert.ToString(Price);	
			string AppMtAmount = Convert.ToString(MortgAmount);	
			string ContactNbr = Convert.ToString(ContactNbr1);	
			string Tel1 = ContactNbr.Substring(0,3);
			string Tel2 = ContactNbr.Substring(3,3);
			string Tel3 = Convert.ToString(ContactNbr2);
			varPostalCode = PostalCode;
			varCreationDate = System.DateTime.Today.ToString("yyyy-MM-dd");
			varRefNbr = AppRefNbr;
			
			
			string cofDate = System.DateTime.Today.AddDays(5).ToString("yyyy-MM-dd");
			string varPropertyAddress = streetNum + " " + varAddressName + " " + varAddressType;
						
			//Order Request
			repo.DomNasHome.RequestService.Click();
						
			//Select Lender (Broker Only)
			repo.DomNasHome.RequestTable.LenderSelection.Click();
			Delay.Milliseconds(100);
			repo.DomNasHome.RequestTable.LenderSelection.TagValue = varLender;
							
			repo.DomNasHome.RequestTable.AppPostalCode.PressKeys(varPostalCode);
			repo.DomNasHome.RequestTable.AppPostalCode.PressKeys("{Tab}");
			Delay.Milliseconds(100);
			
			repo.DomNasHome.RequestTable.AppStreetNum.PressKeys(streetNum);
			repo.DomNasHome.RequestTable.AppStreetName.PressKeys(varAddressName);
			repo.DomNasHome.RequestTable.AppStreetType.PressKeys(varAddressType);
			Delay.Milliseconds(100);
			
			string resReq = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//div[#'serviceTypeDiv']//input[@name='service_type' and @value='$']".Replace("$",varSerType).Trim();
			Ranorex.InputTag serviceSel = resReq;
			serviceSel.Click();
			Delay.Milliseconds(200);
			
			//Lender Email Address
			if (varLender != "Lendwise" && varLender != "TD Broker Services")
			{
				//INput Lender email address
				repo.DomNasHome.MenuDisplay.LenderEmailAddress.Click();
				repo.DomNasHome.MenuDisplay.LenderEmailAddress.PressKeys("an.zhou@nas.com");
			}
			
			//Loan type selected
			repo.DomNasHome.MenuDisplay.Purchase.Click();
			repo.DomNasHome.MenuDisplay.Price.PressKeys(AppPrice);
			repo.DomNasHome.MenuDisplay.MortgageAmt.PressKeys(AppMtAmount);
			
			varPrice = AppPrice;
			varMortgage = AppMtAmount;
			
			//COF field
			if(repo.DomNasHome.MenuDisplay.CofDeadlineInfo.Exists())
			{
			repo.DomNasHome.MenuDisplay.CofDeadline.Click();
			repo.DomNasHome.MenuDisplay.CofDeadline.Value = cofDate;
			}
			
			repo.DomNasHome.MenuDisplay.AppRefNbr.PressKeys(AppRefNbr);
			repo.DomNasHome.MenuDisplay.AppName.PressKeys("Teck Smith");
			repo.DomNasHome.MenuDisplay.AppBraNbr.PressKeys("36888");
			
			
			if (varSerType != "Drive-By" && varSerType != "Desktop")
			{
			repo.DomNasHome.MenuDisplay.AppContactName.PressKeys("John Smith");
			repo.DomNasHome.MenuDisplay.Tel1.PressKeys(Tel1);
			repo.DomNasHome.MenuDisplay.Tel2.PressKeys(Tel2);
			repo.DomNasHome.MenuDisplay.Tel3.PressKeys(Tel3);
			}
						
			//Input additional Email
			//repo.DomNasHome.MenuDisplay.AdditionalEmailBroker.PressKeys("admintest@nationwideappraisals.com");
			
			repo.DomNasHome.MenuDisplay.SpDirection.PressKeys("This is a test, please do not proceed; Thanks");
			
			//Select Payment type
			repo.DomNasHome.MenuDisplay.ApplicantPay.Click();
			
			if (repo.DomNasHome.MenuDisplay.CodEmail != null)
			{
				repo.DomNasHome.MenuDisplay.CodEmail.PressKeys("testsupport@nas.com");
			}
			Delay.Milliseconds(100);
			
			repo.DomNasHome.MenuDisplay.SubmitBtn.Click();
			Delay.Milliseconds(200);
			
			
			//Submit Request
			repo.DomNasHome.MenuDisplay.SelectAppraiserSubmitBtn.Click();
			Delay.Milliseconds(200);
			
			//Get Nas Number
			var NasNbr = repo.DomNasHome.MenuDisplay.ViewAppraisalNumssStrongTag.InnerText.Trim();
			NasNbr = NasNbr.Substring(30, 7);
			varNasNbr = NasNbr;
			
			//Get Service Type from DB
			string SQL = "SELECT a.service_residential FROM app_request a where app_request_nbr  = " + NasNbr + ";";
			DataTable dt = DBquery.RunQuery(SQL);
			string reqSerType = dt.Rows[0][0].ToString().Trim();
			
			
			//Report request submit status
			Report.Log(ReportLevel.Success, "Validation", "Request has been successfully submitted");
			Validate.Exists(repo.DomNasHome.MenuDisplay.H1TagYourRequestHasBeenSubmitted);
			Report.Log(ReportLevel.Info, "Validation", "Request Number is: " + NasNbr);
			Report.Info("Information","Service type is: " + reqSerType);
			Delay.Milliseconds(200);
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
			
		}
	}
}
