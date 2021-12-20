/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 04/02/2016
 * Time: 12:12 PM
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

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace NasAdmin
{
	/// <summary>
	/// Description of AdminSubmitRequest.
	/// Scenario: Nas admin log in and create request based on the selected service type, and uplaod supporting document
	/// </summary>
	[TestModule("774E8A13-4187-4B1B-8284-8DBF4FF21B56", ModuleType.UserCode, 1)]
	public class AdminSubmitRequest : ITestModule
	{
		/// <summary>
		/// Using the NasAdminRepository repository.
		/// </summary>
		public static NasAdminRepository repo = NasAdminRepository.Instance;
		
		static AdminSubmitRequest instance = new AdminSubmitRequest();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public AdminSubmitRequest()
		{
			// Do not delete - a parameterless constructor is required!
		}
		public void KillBrowser(string IE)
		{
		}

		public void ClearBrowserCookies(string IE)
		{
		}
		
		
		#region Varilables
		
		string _varAdmin = "";
		[TestVariable("E1F8CD8E-B3AA-4BEE-A055-40FA97023A6D")]
		public string varAdmin
		{
			get { return _varAdmin; }
			set { _varAdmin = value; }
		}
		
		string _varPwd = "";
		[TestVariable("386F24E8-1032-4DDB-8A65-D46FF6527B3B")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varUrl = "";
		[TestVariable("01448979-5532-4F0F-8FEF-37C8F3A41839")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varPostalCode = "";
		[TestVariable("2CD11C65-1767-4210-B4B5-2F550AFCB609")]
		public string varPostalCode
		{
			get { return _varPostalCode; }
			set { _varPostalCode = value; }
		}
		
		string _varRefNbr = "";
		[TestVariable("E2C91CCC-3AF2-470B-BE0B-ED2EC6AEC5A8")]
		public string varRefNbr
		{
			get { return _varRefNbr; }
			set { _varRefNbr = value; }
		}
		
		string _varPrice = "";
		[TestVariable("7A928520-E8B7-49CF-BB1F-E8FAD02D21ED")]
		public string varPrice
		{
			get { return _varPrice; }
			set { _varPrice = value; }
		}
		
		string _varMorgAmount = "";
		[TestVariable("182F4F59-2F15-4796-B4A0-5BAF7DF97243")]
		public string varMorgAmount
		{
			get { return _varMorgAmount; }
			set { _varMorgAmount = value; }
		}
		
		string _varPropertyAddress = "";
		[TestVariable("0E730B89-AA05-49D7-87BD-1EB36119BA3E")]
		public string varPropertyAddress
		{
			get { return _varPropertyAddress; }
			set { _varPropertyAddress = value; }
		}
		
		string _varSerType = "";
		[TestVariable("732DAC50-FEE1-4A6B-AD8F-84B1BCAAF9BF")]
		public string varSerType
		{
			get { return _varSerType; }
			set { _varSerType = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("504DD898-B4F8-4E86-B8D5-97F4A66AAE52")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varCreateDate = "";
		[TestVariable("489CB192-C3E1-48FC-B16B-5BF5E15ECAF8")]
		public string varCreateDate
		{
			get { return _varCreateDate; }
			set { _varCreateDate = value; }
		}
		
		string _varAppName = "";
		[TestVariable("9AD19044-6EF6-42B9-BDCA-3C594D45AD1C")]
		public string varAppName
		{
			get { return _varAppName; }
			set { _varAppName = value; }
		}
		
		string _varContactName = "";
		[TestVariable("AB5313A4-1D70-4496-8969-FCE5B1CCFFE4")]
		public string varContactName
		{
			get { return _varContactName; }
			set { _varContactName = value; }
		}
		
		string _varBranchNbr = "";
		[TestVariable("6D57E406-D422-44A7-BFFD-0FAAB987A00B")]
		public string varBranchNbr
		{
			get { return _varBranchNbr; }
			set { _varBranchNbr = value; }
		}
		
		string _varContactTel = "";
		[TestVariable("2690ADC5-78CA-4440-A80A-C22AB6FDBD96")]
		public string varContactTel
		{
			get { return _varContactTel; }
			set { _varContactTel = value; }
		}
		
		string _varSpText = "";
		[TestVariable("44C6E0F2-2CE9-42A9-88F6-0FA502E1DF88")]
		public string varSpText
		{
			get { return _varSpText; }
			set { _varSpText = value; }
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
			Delay.Milliseconds(300);
			
			//Admin User log in
			repo.Dom_NasHome.Userid.PressKeys(varAdmin);
			repo.Dom_NasHome.Password.PressKeys(varPwd);
			repo.Dom_NasHome.RememberUserName.Click();
			repo.Dom_NasHome.LoginSubmit.Click();
			Delay.Milliseconds(300);
			
			repo.NotForThisSite.Click();
			
			WebDocument DomWeb = repo.Dom_NasHome.Self; 
			
			//Request Service
			repo.Dom_NasHome.NasAdminMenu.RequestService.Click();
			Delay.Milliseconds(0);
			
			//Generate Random Number
			var random = new Random();
			var StrNbr = random.Next(1, 900);
			var ClientRef = random.Next(123456789, 987654321);
			var Price = random.Next(500,900)*1000;
			var MortgAmount= random.Next(100,500)*1000;
			var ContactNbr1 = random.Next(111111, 999999);
			var ContactNbr2 = random.Next(1111, 9999);
			
			const string varAddressName = "Ranorex Insanity";   //"Ranorex Insanity";
			const string varAddressType = "Dr";
				
			string streetNum = Convert.ToString(StrNbr);
			string AppRefNbr = Convert.ToString(ClientRef);	
			string AppPrice = Convert.ToString(Price);	
			string AppMtAmount = Convert.ToString(MortgAmount);	
			string ContactNbr = Convert.ToString(ContactNbr1);	
			string Tel1 = ContactNbr.Substring(0,3);
			string Tel2 = ContactNbr.Substring(3,3);
			string Tel3 = Convert.ToString(ContactNbr2);
			
			varPropertyAddress = streetNum + " " + varAddressName + " " + varAddressType;
			varRefNbr = AppRefNbr;
			varSpText = "This is a test, please do not proceed; Thanks";
			
			//Fill in Order information
			repo.Dom_NasHome.MenuDisplay.RequestService.AppPostalCode.PressKeys(varPostalCode);
			repo.Dom_NasHome.MenuDisplay.RequestService.AppPostalCode.PressKeys("{Tab}");
			Delay.Milliseconds(100);
			
			repo.Dom_NasHome.MenuDisplay.RequestService.AppStreetNum.PressKeys(streetNum);
			repo.Dom_NasHome.MenuDisplay.RequestService.AppStreetName.PressKeys(varAddressName);
			repo.Dom_NasHome.MenuDisplay.RequestService.AppStreetType.PressKeys(varAddressType);
			Delay.Milliseconds(100);
			
			string findStr = "/dom[@domain='test.nationwideappraisals.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//div[#'serviceTypeDiv']//input[@type='radio' and @value~'$']";
			
			string replStr = findStr.Replace("$", varSerType);
			
			//Select Service Type
			InputTag serviceSel = replStr;
			serviceSel.Click();
			Delay.Milliseconds(100);
			
			repo.Dom_NasHome.MenuDisplay.RequestService.PurchaseRadio.Click();
			repo.Dom_NasHome.MenuDisplay.RequestService.Price.PressKeys(AppPrice);
			repo.Dom_NasHome.MenuDisplay.RequestService.MortgageAmt.PressKeys(AppMtAmount);
			
			varPrice = "$" + AppPrice + ".0";
			varMorgAmount = "$" + AppMtAmount + ".0";
			
			varAppName = "John Doe";
			varContactName = "Terry Doe";
			varBranchNbr = "36888";
			
			repo.Dom_NasHome.MenuDisplay.RequestService.AppRefNbr.PressKeys(AppRefNbr);
			repo.Dom_NasHome.MenuDisplay.RequestService.AppBraNbr.PressKeys(varBranchNbr);
			repo.Dom_NasHome.MenuDisplay.RequestService.AppName.PressKeys(varAppName);
			
			repo.Dom_NasHome.MenuDisplay.RequestService.AppContactName.PressKeys(varContactName);
			repo.Dom_NasHome.MenuDisplay.RequestService.Tel1.PressKeys(Tel1);
			repo.Dom_NasHome.MenuDisplay.RequestService.Tel2.PressKeys(Tel2);
			repo.Dom_NasHome.MenuDisplay.RequestService.Tel3.PressKeys(Tel3);
			
			repo.Dom_NasHome.MenuDisplay.RequestService.SpDirectionText.PressKeys(varSpText);
			repo.Dom_NasHome.MenuDisplay.RequestService.InvoiceRadio.Click();
			repo.Dom_NasHome.MenuDisplay.RequestService.SubmitRequestBtn.Click();
			Delay.Milliseconds(100);
			
			repo.Dom_NasHome.MenuDisplay.RequestService.SelAppraiserSubmitBtn.Click();
			Delay.Milliseconds(200);
			
			//Get Nas Number
			var NasNbr = repo.Dom_NasHome.MenuDisplay.RequestService.StrongTagViewNasNbr.InnerText.Trim();
			NasNbr = NasNbr.Substring(30, 7);
			varNasNbr = NasNbr;
			varContactTel = Tel1 + "-" + Tel2 + "-" + Tel3;
			
			//Get Request Creation Date
			varCreateDate = System.DateTime.Today.ToString("yyyy-MM-dd");
			
			
			//Report request submit status
			Report.Log(ReportLevel.Success, "Validation", "Request has been successful submit.");
			Report.Log(ReportLevel.Info, "Validation", "Request Number is: " + NasNbr);
			Validate.Exists(repo.Dom_NasHome.MenuDisplay.RequestService.StrongTagYourAppraisalHasBeenSen);
			Delay.Milliseconds(200);
			
			/* UPOALD SUPPORTMENT DOCUMENTS*/
			repo.Dom_NasHome.MenuDisplay.RequestService.StrongTagUploadSupporting.Click();
			Delay.Milliseconds(100);
			
			repo.Dom_NasHome.MenuDisplay.RequestService.File.DoubleClick();
			Delay.Milliseconds(100);
			
			string filePath = @"c:\Upload_Files\SupportDocument.pdf\";
			
			repo.ChooseFileToUpload.FilePath.PressKeys(filePath);
			Keyboard.Press("{Return}");
			Delay.Milliseconds(300);
			
			repo.Dom_NasHome.MenuDisplay.RequestService.UploadBtn.Click();
			Delay.Milliseconds(500);
			
			//Report Uplaod support document status
			Report.Log(ReportLevel.Success, "Validation", "Uploading Support document successful.");
			Validate.Exists(repo.Dom_NasHome.MenuDisplay.RequestService.SupportingDocumentsHaveBeenUploaded);
			Delay.Milliseconds(300);
			
			//Close Browser
			//Host.Local.KillBrowser("IE");
			//Delay.Milliseconds(0);
			
		}
	}
}
