/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-05-05
 * Time: 11:48 AM
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
	/// Description of RequestService.
	/// </summary>
	[TestModule("0ED33176-AB8F-4069-AC90-3EB614ECFB04", ModuleType.UserCode, 1)]
	public class RequestService : ITestModule
	{
		
		public static Scotia_PortalRepository repo = Scotia_PortalRepository.Instance;
		
		static RequestService instance = new RequestService();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public RequestService()
		{
			// Do not delete - a parameterless constructor is required!
		}
		
		#region Varilables
		
		string _varUid = "";
		[TestVariable("75600886-1E69-4D12-8598-46861AE6E590")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("FC2A1167-503F-45BF-A1E2-C3E69B0BDB0B")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varServiceType = "";
		[TestVariable("A173E8B5-234B-4D93-9527-94B47537B2C0")]
		public string varServiceType
		{
			get { return _varServiceType; }
			set { _varServiceType = value; }
		}
		
		string _varLoanType = "";
		[TestVariable("C1B363ED-EE54-4EA6-9C03-DD49D7235F0F")]
		public string varLoanType
		{
			get { return _varLoanType; }
			set { _varLoanType = value; }
		}
		
		string _varPoCode = "";
		[TestVariable("E1587EBF-8A32-4564-BE21-9FE289341A8A")]
		public string varPoCode
		{
			get { return _varPoCode; }
			set { _varPoCode = value; }
		}
	
		string _varProperty = "";
		[TestVariable("3768EFD5-5FD0-4B6C-807A-4C07B1E5A012")]
		public string varProperty
		{
			get { return _varProperty; }
			set { _varProperty = value; }
		}
		
		string _varServiceVal = "";
		[TestVariable("0809C924-F366-48E3-829F-A07D8C8EAF0F")]
		public string varServiceVal
		{
			get { return _varServiceVal; }
			set { _varServiceVal = value; }
		}
		
		string _varSAoption = "";
		[TestVariable("D31E369A-BA00-4CD7-804F-FE1C6654FCA2")]
		public string varSAoption
		{
			get { return _varSAoption; }
			set { _varSAoption = value; }
		}
		
		string _varOther = "";
		[TestVariable("44529C6C-96DB-4D62-AE05-7099CB963896")]
		public string varOther
		{
			get { return _varOther; }
			set { _varOther = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("BF326F69-143D-4BBC-A292-A8D6231A6B96")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		
		string _varIVP_ServiceType = "";
		[TestVariable("3990501A-80F9-48AA-9846-4D83AE539AC3")]
		public string varIVP_ServiceType
		{
			get { return _varIVP_ServiceType; }
			set { _varIVP_ServiceType = value; }
		}
		
		#endregion
		
		private const string addiEmail = "testsupport@nas.com";
		private const string price = "860000";
		private const string morgAmount = "560000";
		private const string appName = "Ted Smith";
		private const string contact = "Donald Dump";
		private const string spText = "Test only";
		private const string strName = "Scotia Insanity test";
		private const string strType = "Dr"; 
		
		
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		/// 
		public void selectResidentialRequest(string serviceType)
		{
			switch (serviceType)
				{
    				case "Appraisal":
						repo.DomScotia.RequestService.ResidentialRequestType.Appraisal.Click();
       				    break;
    				case "Site Inspection - Mobile - Mortgage":
       				    repo.DomScotia.RequestService.ResidentialRequestType.SiteInspectionMobileMortgage.Click();
       				    repo.MessageFromWebpage.ButtonOK.Click();
        			    break;
        			case "Site Inspection - Modular - Mortgage":
       				    repo.DomScotia.RequestService.ResidentialRequestType.SiteInspectionModularMortgage.Click();
       				    repo.MessageFromWebpage.ButtonOK.Click();
        			    break;
        			case "Site Inspection - Mobile - SPL":
       				    repo.DomScotia.RequestService.ResidentialRequestType.SiteInspectionMobileSPL.Click();
       				    repo.MessageFromWebpage.ButtonOK.Click();
        			    break;   
        			case "Site Inspection - Modular - SPL":
       				    repo.DomScotia.RequestService.ResidentialRequestType.SiteInspectionModularSPL.Click();
       				    repo.MessageFromWebpage.ButtonOK.Click();
        			    break;       
        			case "Stand Alone Rental Schedule (A)":
       				    repo.DomScotia.RequestService.ResidentialRequestType.StandAloneRentalScheduleA.Click();
       				    repo.MessageFromWebpage.ButtonOK.Click();
        			    break;  
        			case "Capitalization":
       				    repo.DomScotia.RequestService.ResidentialRequestType.Capitalization.Click();
        			    break;  
        			case "Progress Advance":
        			    repo.DomScotia.RequestService.ResidentialRequestType.ProgressAdvance.Click();
        			    repo.DomScotia.RequestService.ResidentialRequestType.InspectionN.TagValue = "1";                  //Select Progress Advance Inpection Number
        			    break;  
    				default:
        			    MessageBox.Show("Please specify residential service type.");
        			break;
				}
		}    //End of [selectResidentialRequest] function
		
		public void requestInfo(string poCode, string stNbr, string stName, string stType, string service, string loan, 
		                        string purchasePrice, string mortAmount, string propertyType, string serValue, 
		                        string saOption, string other, string refNbr, string appName, string contact,
		                        string tel1, string tel2, string tel3, string spTxt, string email)
		{
			//Property Address
			repo.DomScotia.RequestService.PropertyAddress.AppPostalCode.PressKeys(poCode);
			repo.DomScotia.RequestService.PropertyAddress.AppPostalCode.PressKeys("{Tab}");
			Delay.Milliseconds(100);
			
			repo.DomScotia.RequestService.PropertyAddress.AppStreetNum.PressKeys(stNbr);
			repo.DomScotia.RequestService.PropertyAddress.AppStreetName.PressKeys(stName);
			repo.DomScotia.RequestService.PropertyAddress.AppStreetType.PressKeys(stType);
			Delay.Milliseconds(100);
			
			if(service != "Capitalization" && service != "Progress Advance")
			{//Loan Purpose
			string loanPath = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[#'$']".Replace("$", loan).Trim();
			Ranorex.InputTag loanOption = loanPath;
			loanOption.Click();
			Delay.Milliseconds(100);
			
			repo.DomScotia.RequestService.LoanPurpose.Price.PressKeys(purchasePrice);
			repo.DomScotia.RequestService.LoanPurpose.MortgageAmt.PressKeys(mortAmount);
			}
			
			//Property type Options
			//Fill out property type and service type section for 'Appraisal' request
			if (service == "Appraisal")
			{
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
				int j= 5;
				if (loan == "Refinance" || loan == "TransferSwitchwithAdditionalFunds")
				{ j= 6;}
				for (int i = 1; i<=j; i++)
				{
				string x = i.ToString() + other;
				Ranorex.InputTag otherSel = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[#'other_inform$']".Replace("$",x).Trim();
				otherSel.Click();
				}
			} //End of Property type and other section
			
			
			//Loan and Contact Information
			repo.DomScotia.RequestService.LoanAndContactInformation.AppRefNbr.PressKeys(refNbr);
			repo.DomScotia.RequestService.LoanAndContactInformation.AppName.PressKeys(appName);
			
			//No contact name for "Captalization"
			if(service != "Capitalization")
			{
			repo.DomScotia.RequestService.LoanAndContactInformation.AppContactName.PressKeys(contact);
			repo.DomScotia.RequestService.LoanAndContactInformation.Tel1.PressKeys(tel1);
			repo.DomScotia.RequestService.LoanAndContactInformation.Tel2.PressKeys(tel2);
			repo.DomScotia.RequestService.LoanAndContactInformation.Tel3.PressKeys(tel3);
			repo.DomScotia.RequestService.LoanAndContactInformation.PreferNumber1.Click();
			Delay.Milliseconds(100);
			}
			//End of Loan and Contact Information section
			
			//Special Instructions
			repo.DomScotia.RequestService.SpecialInstructions.SpText.PressKeys(spTxt);
			
			//Email notifications
			repo.DomScotia.RequestService.EmailNotifications.AdditionalEmail.PressKeys(email);
			Delay.Milliseconds(100);
			
			//Select Payment type as "Invoice"   (Added @ 2016.05.16)
			if(repo.DomScotia.RequestService.Payment.InvoiceInfo.Exists())
			{repo.DomScotia.RequestService.Payment.Invoice.Click();}
			
			//Submit request
			repo.DomScotia.RequestService.SubmitRequest.MoveTo();
			Delay.Milliseconds(100);
			repo.DomScotia.RequestService.SubmitRequest.PerformClick();
			Delay.Milliseconds(500);
			
			//Report order submit status
			Validate.Exists(repo.DomScotia.H1TagYourRequestHasBeenSubmitted);
			Validate.Exists(repo.DomScotia.PTagAppraisalHasBeenSentToAppraiser);
			Report.Log(ReportLevel.Success, "Validation", "Request has been successfullu submitted.");
			
		}   //End of [requestInfo] function
		
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
		
		
		//===================================================================================
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
			
			repo.DomScotia.RequestService.ResidentialNewRequest.Click();
			
			selectResidentialRequest(varServiceType);
			
			//Fill in request information in request page
			requestInfo(varPoCode, stNbr, strName, strType, varServiceType, varLoanType, price, morgAmount, varProperty, 
			            varServiceVal, varSAoption, varOther, clientRef, appName, contact, 
			            contactNbr1, contactNbr2, contactNbr3, spText, addiEmail);
			
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
