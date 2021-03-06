/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-05-09
 * Time: 1:59 PM
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

namespace Scotia_Portal
{
	/// <summary>
	/// Description of ScotiaDesktop.
	/// </summary>
	[TestModule("611BDF9F-6E41-42B8-8F3A-5756FCF38B21", ModuleType.UserCode, 1)]
	public class ScotiaDesktop : ITestModule
	{
		public static Scotia_PortalRepository repo = Scotia_PortalRepository.Instance;
		
		static ScotiaDesktop instance = new ScotiaDesktop();
		
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ScotiaDesktop()
		{
			// Do not delete - a parameterless constructor is required!
		}
			
		#region Varilables
		
		string _varUid = "";
		[TestVariable("4B799192-28CF-487A-9D62-69B561466A9C")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
				
		string _varServiceType = "";
		[TestVariable("25B3491F-A23B-44CE-96FA-4F115C617CA7")]
		public string varServiceType
		{
			get { return _varServiceType; }
			set { _varServiceType = value; }
		}
		
		string _varPwd = "";
		[TestVariable("B3BBDC06-528B-4C85-AD03-208FACE4B590")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
	
		string _varLoanType = "";
		[TestVariable("CE7E42E8-8DAC-46FE-8D99-E47C5ADE4307")]
		public string varLoanType
		{
			get { return _varLoanType; }
			set { _varLoanType = value; }
		}
		
		string _varPoCode = "";
		[TestVariable("E1A43C97-1F60-4580-90AE-EEB7FAADAE43")]
		public string varPoCode
		{
			get { return _varPoCode; }
			set { _varPoCode = value; }
		}
		
		string _varProperty = "";
		[TestVariable("959FAF77-302D-4333-9667-115B0EE6D900")]
		public string varProperty
		{
			get { return _varProperty; }
			set { _varProperty = value; }
		}
		
		string _varServiceVal = "";
		[TestVariable("BAEB8223-A005-4AB3-9DD3-FEB5F8CB2DD1")]
		public string varServiceVal
		{
			get { return _varServiceVal; }
			set { _varServiceVal = value; }
		}
		
		string _varSAoption = "";
		[TestVariable("0FD68BA3-4A67-4073-837D-B315FED95035")]
		public string varSAoption
		{
			get { return _varSAoption; }
			set { _varSAoption = value; }
		}
		
		string _varOther = "";
		[TestVariable("538904DC-F731-4C93-A1A8-8F51B765EE7A")]
		public string varOther
		{
			get { return _varOther; }
			set { _varOther = value; }
		}
		
		string _varOwn12Mths = "";
		[TestVariable("A7A839BF-4175-49D4-9636-0E65FA409202")]
		public string varOwn12Mths
		{
			get { return _varOwn12Mths; }
			set { _varOwn12Mths = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("D1782C22-6005-4220-91D3-4752E4198DAB")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varIVP_serviceType = "";
		[TestVariable("AF655D31-784B-447E-B8C7-1CA40C6F3AFE")]
		public string varIVP_serviceType
		{
			get { return _varIVP_serviceType; }
			set { _varIVP_serviceType = value; }
		}
		
		
		string _varPrice = "";
		[TestVariable("2DDC243C-0B81-493A-A491-609F5D43DC64")]
		public string varPrice
		{
			get { return _varPrice; }
			set { _varPrice = value; }
		}
		
		string _varMortgage = "";
		[TestVariable("DF52D2AD-9248-49B8-BD34-308CC4D217AA")]
		public string varMortgage
		{
			get { return _varMortgage; }
			set { _varMortgage = value; }
		}
		
		#endregion
		
		private const string addiEmail = "supportuattest.nas.com";
		private const string appName = "Holly Molly";
		private const string contact = "Johhny Sack";
		private const string spText = "Scotia Desktop test.";
		private const string strName = "Front end";
		private const string strType = "Road"; 
		
		
		//==================================================================================
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
		
		public void OrderDesktop (string service, string poCode, string stNbr, string stName, string stType, string loan, string purchasePrice, string mortAmount,
		                        string propertyType, string serValue, string saOption, string other, string own, string refNbr, string appName, string contact, 
		                        string tel1, string tel2, string tel3, string spTxt, string email)
		{
			repo.DomScotia.RequestService.ResidentialNewRequest.Click();
			Delay.Milliseconds(100);
			
			//Select 'Appraisal' request Type
			if(service !="Appraisal")
			{MessageBox.Show("Please check your request type is 'Appraisal'.");}else
			{
				repo.DomScotia.RequestService.ResidentialRequestType.Appraisal.Click();
				Delay.Milliseconds(100);
			}
       		
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
			
			repo.DomScotia.RequestService.LoanPurpose.Price.PressKeys(purchasePrice);
			repo.DomScotia.RequestService.LoanPurpose.MortgageAmt.PressKeys(mortAmount);
			
			
			//Fill out property type and service type section for 'Appraisal' request
			
			//Property type Options
			string proptyOptionPath = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[#'propertyType_$']".Replace("$", propertyType).Trim();
			Ranorex.InputTag proptyOption = proptyOptionPath;
			proptyOption.Click();
			Delay.Milliseconds(200);
			
			//Service Type Options
			Ranorex.InputTag serviceValue = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[#'$']".Replace("$", serValue).Trim();
			serviceValue.MoveTo();
			Delay.Milliseconds(100);
			serviceValue.Click();
					if(serValue != "AsisValueOnly")
					{ repo.MessageFromWebpage.ButtonOK.Click();}
			Delay.Milliseconds(100);
			
			//SA and Rush Options
			if ( saOption == "No")
			{ repo.DomScotia.RequestService.ServiceTypeOptions.ScheduleappraisalNo.Click();}
			else if (saOption == "Yes")
			{ 	repo.DomScotia.RequestService.ServiceTypeOptions.ScheduleappraisalYes.Click();}
			  
				repo.DomScotia.RequestService.ServiceTypeOptions.RushserviceNo.Click();
			
			//Other Information
				for (int i = 1; i<=5; i++)
				{
				string x = i.ToString() + other;
				Ranorex.InputTag otherSel = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[#'other_inform$']".Replace("$",x).Trim();
				otherSel.Click();
				}
				//Select Owned for 12 months option
				if (loan == "Refinance" || loan == "TransferSwitchwithAdditionalFunds")
				{ 
				Ranorex.InputTag own12Mths = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[#'other_inform6$']".Replace("$",own).Trim();
				own12Mths.Click();
				}
			//End of Property type and other section
			
			
			//Loan and Contact Information
			repo.DomScotia.RequestService.LoanAndContactInformation.AppRefNbr.PressKeys(refNbr);
			repo.DomScotia.RequestService.LoanAndContactInformation.AppName.PressKeys(appName);
			
			//No contact name for "Captalization"
			repo.DomScotia.RequestService.LoanAndContactInformation.AppContactName.PressKeys(contact);
			repo.DomScotia.RequestService.LoanAndContactInformation.Tel1.PressKeys(tel1);
			repo.DomScotia.RequestService.LoanAndContactInformation.Tel2.PressKeys(tel2);
			repo.DomScotia.RequestService.LoanAndContactInformation.Tel3.PressKeys(tel3);
			repo.DomScotia.RequestService.LoanAndContactInformation.PreferNumber1.Click();
			Delay.Milliseconds(100);
			
			//End of Loan and Contact Information section
			
			//Special Instructions
			repo.DomScotia.RequestService.SpecialInstructions.SpText.PressKeys(spTxt);
						
			/* comment out htis part for non-Scotia BMU to speed up execution
			if (repo.DomScotia.RequestService.Payment.InvoiceInfo.Exists())
			{
				repo.DomScotia.RequestService.Payment.Invoice.Click();
				Delay.Milliseconds(100);
			}
			*/
			
			//Email notifications
			repo.DomScotia.RequestService.EmailNotifications.AdditionalEmail.PressKeys(email);
			Delay.Milliseconds(100);
			
			//Submit request
			repo.DomScotia.RequestService.SubmitRequest.MoveTo();
			Delay.Milliseconds(100);
			repo.DomScotia.RequestService.SubmitRequest.PerformClick();
			Delay.Milliseconds(500);
			
			//Report order submit status
			Validate.Exists(repo.DomScotia.H1TagYourRequestHasBeenSubmitted);
			Report.Log(ReportLevel.Success, "Validation", "Request has been successfully submitted.");
			
		}
				
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		
		
		//=================================================================
		
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			Delay.Milliseconds(200);
			
			var random = new Random();
			string stNbr = random.Next(1, 900).ToString();
			string clientRef = random.Next(123456789, 987654321).ToString();
			//int Price = random.Next(500,900)*1000;
			//int MortgAmount= random.Next(100,500)*1000;
			string contactNbr1 = random.Next(111, 999).ToString();
			string contactNbr2 = random.Next(111, 999).ToString();
			string contactNbr3 = random.Next(1111,9999).ToString();
			
			Login UsrLogin = new Login();
			UsrLogin.lauchScotia();
			UsrLogin.UserLogin(varUid, varPwd);
			
			//Request New Service
			repo.DomScotia.MainMenu.RequestService.Click();
			Delay.Milliseconds(100);
			
			//Order Desktop Request		
			//Fill in request information in request page
			OrderDesktop(varServiceType, varPoCode, stNbr, strName, strType, varLoanType, varPrice, varMortgage, varProperty, 
			            varServiceVal, varSAoption, varOther, varOwn12Mths, clientRef, appName, contact, 
			            contactNbr1, contactNbr2, contactNbr3, spText, addiEmail);
			
			Delay.Milliseconds(100);
			
			//Get Nas number
			varNasNbr = getRequestNumber();
			string reqServiceType = getRequestServiceType();
						
			Report.Log(ReportLevel.Info, varNasNbr + " request service type is: " + reqServiceType);
			Validate.AreEqual(reqServiceType, varIVP_serviceType);
			
			Delay.Milliseconds(200);
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
			
		}
	}
}
