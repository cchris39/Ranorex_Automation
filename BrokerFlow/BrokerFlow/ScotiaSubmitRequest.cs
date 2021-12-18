/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 01/03/2016
 * Time: 12:03 PM
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
using MySql.Data.MySqlClient;
using System.Data;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace BrokerFlow
{
	
	/// <summary>
	/// Description of ScotiaSubmitRequest.
	/// </summary>
	[TestModule("2A8822B9-92C8-4A3B-9C1B-9D48B499334D", ModuleType.UserCode, 1)]
	public class ScotiaSubmitRequest : ITestModule
	{
		/// <summary>
		/// Using the Dom_AppraiserSanityTestRepository repository.
		/// </summary>
		public static BrokerFlowRepository repo = BrokerFlowRepository.Instance;
		
		static ScotiaSubmitRequest instance = new ScotiaSubmitRequest();
		
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("34DB1DA4-AEEC-4310-BAC3-B2D2BA7824A4")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varPwd = "";
		[TestVariable("73A84531-BA52-4CAD-BF41-217A3873538D")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varUid = "";
		[TestVariable("CE996F4A-D240-4CDD-BC74-1D2F27493A64")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varUrl = "";
		[TestVariable("74CBBA9C-4078-4015-BA2B-0F9A98DE643A")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varPostalCode = "";
		[TestVariable("B170EE97-FD1C-49BE-A997-B34B5825E19F")]
		public string varPostalCode
		{
			get { return _varPostalCode; }
			set { _varPostalCode = value; }
		}
		
		string _varPrice = "";
		[TestVariable("ABFF4D7D-6CC7-4A81-BD03-F8F81A4A202F")]
		public string varPrice
		{
			get { return _varPrice; }
			set { _varPrice = value; }
		}
		
		string _varMortgage = "";
		[TestVariable("1721A5A5-D24C-409B-857A-E920E0CB9FB0")]
		public string varMortgage
		{
			get { return _varMortgage; }
			set { _varMortgage = value; }
		}
		
		string _varPropertyAddress = "";
		[TestVariable("ABE13A40-F221-4D82-861D-01FA94942E41")]
		public string varPropertyAddress
		{
			get { return _varPropertyAddress; }
			set { _varPropertyAddress = value; }
		}
		
		string _varRefNbr = "";
		[TestVariable("4EB50B31-F6A5-485F-B11A-5248B529D13F")]
		public string varRefNbr
		{
			get { return _varRefNbr; }
			set { _varRefNbr = value; }
		}
		
		string _varCreationDate = "";
		[TestVariable("9CD86ED0-E9D6-40E2-BD44-54A9C511409A")]
		public string varCreationDate
		{
			get { return _varCreationDate; }
			set { _varCreationDate = value; }
		}
		
		string _varLender = "";
		[TestVariable("98CAEEFC-4E36-478E-BC66-9ECC6A24363B")]
		public string varLender
		{
			get { return _varLender; }
			set { _varLender = value; }
		}
		
		string _resReqType = "";
		[TestVariable("A3F5C3A7-0930-4F7E-929E-3DEC874ADE36")]
		public string resReqType
		{
			get { return _resReqType; }
			set { _resReqType = value; }
		}
		
		string _loanType = "";
		[TestVariable("3670E10E-8926-4E31-8A36-7F548C92BCBF")]
		public string loanType
		{
			get { return _loanType; }
			set { _loanType = value; }
		}
		
		string _propertyType = "";
		[TestVariable("37359FE9-D59F-4930-AA1C-06CD5BE07C2A")]
		public string propertyType
		{
			get { return _propertyType; }
			set { _propertyType = value; }
		}
		
		string _serValOption = "";
		[TestVariable("ADDB4D3C-BB82-44FF-A17A-F49D3C1B0B4E")]
		public string serValOption
		{
			get { return _serValOption; }
			set { _serValOption = value; }
		}
		
		string _otherInfo = "";
		[TestVariable("1DFEB3D5-471C-4B2F-A98A-D0FF797998E4")]
		public string otherInfo
		{
			get { return _otherInfo; }
			set { _otherInfo = value; }
		}
		
		string _varOwnOption = "";
		[TestVariable("62309BBF-EB76-4FAF-878D-235AEE02DA32")]
		public string varOwnOption
		{
			get { return _varOwnOption; }
			set { _varOwnOption = value; }
		}
		
		string _varSAoption = "";
		[TestVariable("C68D061D-E308-4764-988C-5BC8DA24F177")]
		public string varSAoption
		{
			get { return _varSAoption; }
			set { _varSAoption = value; }
		}
		
		string _varAsgAppraiser = "";
		[TestVariable("ABB1EF14-56E5-4E0B-AA5C-DB0CD6E320F7")]
		public string varAsgAppraiser
		{
			get { return _varAsgAppraiser; }
			set { _varAsgAppraiser = value; }
		}
		
		string _postCode = "";
		[TestVariable("852CA5FB-54AC-46FC-B8F4-D3DAB6B818D3")]
		public string postCode
		{
			get { return _postCode; }
			set { _postCode = value; }
		}
		
		#endregion
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ScotiaSubmitRequest()
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
			Delay.Milliseconds(200);
			
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
			
			const string varAddressName = "Scotia BRM";   //"Ranorex Insanity";
			const string varAddressType = "Broker Test";
				
			string streetNum = Convert.ToString(StrNbr);
			string AppRefNbr = Convert.ToString(ClientRef);	
			string AppPrice = Convert.ToString(Price);	
			string AppMtAmount = Convert.ToString(MortgAmount);	
			string ContactNbr = Convert.ToString(ContactNbr1);	
			string Tel1 = ContactNbr.Substring(0,3);
			string Tel2 = ContactNbr.Substring(3,3);
			string Tel3 = Convert.ToString(ContactNbr2);
			
			string varPropertyAddress = streetNum + " " + varAddressName + " " + varAddressType;
			varPostalCode = postCode;
			varCreationDate = System.DateTime.Today.ToString("yyyy-MM-dd");
			varRefNbr = AppRefNbr;
			
			
			//Get Nas number which is completed for ordering PI request if there is any
			string SQL1 = "SELECT a.`app_request_nbr` FROM app_request a where lender = '" + varLender + "' " + "and status = 'Completed' and service_residential = 'Full-Service' order by app_request_nbr desc limit 1;";
			DataTable dt1 = DBquery.RunQuery(SQL1);
			string oriPI = dt1.Rows[0][0].ToString().Trim();
			
			//=======================================================================================================================
			//Order Request
			repo.DomNasHome.RequestService.Click();
						
			//Select Lender (Broker Only)
			repo.DomNasHome.RequestTable.LenderSelection.Click();
			repo.DomNasHome.RequestTable.LenderSelection.TagValue = varLender;
			Delay.Milliseconds(100);
			
			repo.DomNasHome.ScotiaBRM.AppraisalServiceType.Click();
			Delay.Milliseconds(100);
			
			string resReq = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//table[#'table_new_request_1']//input[#'$']".Replace("$", resReqType);
			string loanPurpose = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//table[#'table_loan_purpose']//input[#'$']".Replace("$", loanType);
			string propertyOption = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[@id~'propertyType_*' and @tagvalue='$']".Replace("$", propertyType);
			string serValueOp = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[#'$']";
			
			Ranorex.InputTag resReqOption = resReq;
			Ranorex.InputTag loanOpn = loanPurpose; 
			Ranorex.InputTag serOption = serValueOp.Replace("$", serValOption);
			Ranorex.InputTag propOption = propertyOption;
			resReqOption.Click();
			Delay.Milliseconds(100);
				
				if (resReqType == "ProgressInspection")
				{
					repo.DomNasHome.ScotiaBRM.InspectionN.TagValue = "1";
					repo.DomNasHome.ScotiaBRM.OrgReqNbrPA.PressKeys(oriPI);
					repo.DomNasHome.ScotiaBRM.SearchOriNbrBtn.Click();
					Delay.Milliseconds(200);
					loanOpn.Click();
					repo.DomNasHome.RequestTable.AdditionalEmail.PressKeys("an.zhou@nas.com");
					
					
				}else
				{
					if (resReqType == "SiteInspection-Mobile-Mortgage" || resReqType == "SiteInspection-Modular-Mortgage" ||  resReqType == "StandAloneRentalScheduleA")
					{
					//Acknowledge Pop Up message
					repo.MessageFromWebpage.ButtonOK.Click();
					}
					//Select Loan Purpose Option
					//repo.DomNasHome.ScotiaBRM.LoanPurpose.TagValue = loanType;
					loanOpn.Click();
					Delay.Milliseconds(200);
			
					if (resReqType != "Progress Inspection")
					{
					repo.DomNasHome.ScotiaBRM.PurchasePrice.PressKeys(AppPrice);
					repo.DomNasHome.ScotiaBRM.LoanAmt.PressKeys(AppMtAmount);
							
					varPrice = AppPrice;
					varMortgage = AppMtAmount;
					}
			
					//Fill in the multiple options for 'Appraisal' service Type			
					if ( resReqType == "Appraisal")
					{
						propOption.Click();
						Delay.Milliseconds(100);
						serOption.Click();
						if (serValOption != "AsisValueOnly")
						{
						repo.MessageFromWebpage.ButtonOK.Click();
						}
						Ranorex.InputTag SAoption = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[@id~'scheduleappraisal$']".Replace("$",varSAoption);
						SAoption.Click();
						Delay.Milliseconds(100);
					
					//Other information
					for (int i = 1; i <= 5; i++)
						{
						string j = i.ToString();
						Ranorex.InputTag otherOption = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[@id>'other_informx$']".Replace("x",j).Replace("$",otherInfo);
						otherOption.Click();
						}
						
						if (loanType == "Refinance" || loanType == "TransferSwitchwithAdditionalFunds")
						{
							//If owned for last 12 mths
							Ranorex.InputTag own12Mths = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[@id>'other_inform6$']".Replace("$",varOwnOption);
							own12Mths.Click();
							Delay.Milliseconds(200);
						}
					}
			
				repo.DomNasHome.RequestTable.AppPostalCode.PressKeys(postCode);
				repo.DomNasHome.RequestTable.AppPostalCode.PressKeys("{Tab}");
				Delay.Milliseconds(100);
			
				repo.DomNasHome.RequestTable.AppStreetNum.PressKeys(streetNum);
				repo.DomNasHome.RequestTable.AppStreetName.PressKeys(varAddressName);
				repo.DomNasHome.RequestTable.AppStreetType.PressKeys(varAddressType);
				Delay.Milliseconds(100);
			
				repo.DomNasHome.RequestTable.AppRefNbr.PressKeys(AppRefNbr);
				//repo.DomNasHome.RequestTable.AppBraNbr.PressKeys("36888");                                       //Remove Branch Number for Broker as per Scotia Request
				repo.DomNasHome.RequestTable.AppName.PressKeys("Teck Smith");
			
				repo.DomNasHome.RequestTable.AppContactName.PressKeys("John Smith");
				repo.DomNasHome.RequestTable.ContactTel1.PressKeys(Tel1);
				repo.DomNasHome.RequestTable.ContactTel2.PressKeys(Tel2);
				repo.DomNasHome.RequestTable.ContactTel3.PressKeys(Tel3);
				repo.DomNasHome.RequestTable.PreferNumber.Click();
			
				repo.DomNasHome.RequestTable.SpecialInstruction.PressKeys("Test only, please do not proceed; Thanks");
				repo.DomNasHome.RequestTable.AdditionalEmail.PressKeys("testsupport@nas.com");
				Delay.Milliseconds(100);
				}
			
			//Choose Payment Type
			repo.DomNasHome.ScotiaBRM.ApplicantToPay.Click();
			repo.DomNasHome.ScotiaBRM.CodEmail.PressKeys("testadmin@nas.com");
			Delay.Milliseconds(100);
			
			//Choose Appraiser Type
			if (varAsgAppraiser == "D")
			{
				//Direct Appraiser
				repo.DomNasHome.RequestTable.DirectAppraiserType.Click();
				repo.DomNasHome.MenuDisplay.SubmitBtn.Click();
				Delay.Milliseconds(500);
			}else if (varAsgAppraiser == "A")
			{
				repo.DomNasHome.RequestTable.AutoAppraiserType.Click();
				repo.DomNasHome.MenuDisplay.SubmitBtn.Click();
				Delay.Milliseconds(300);
			}
			
			repo.DomNasHome.ScotiaBRM.SelectandSubmitBtn.Click();       
			
			//Check if the request is determined as Desktop
			if (repo.DomNasHome.ScotiaBRM.ScotiaDesktopMsgInfo.Exists(200))
			{
				repo.DomNasHome.ScotiaBRM.DesktopOK.Click();             	//Submit Desktop order
				//repo.DomNasHome.ScotiaBRM.DesktopUPGRADE.Click();        //Upgrade Desktop request
				//Report.Log(ReportLevel.Info,"Information", "Desktop request has been upgraded");
			}
			
			//Get Nas Number
			Delay.Milliseconds(200);
			var NasNbr = repo.DomNasHome.ScotiaBRM.StrongTagViewAppraisalNum.InnerText.Trim();            //If select 'Applicant to Pay'
			NasNbr = NasNbr.Substring(30, 7);
			
			//var NasNbr = repo.DomNasHome.ScotiaBRM.FontTagNasNbr.InnerText.Trim();                      //If select 'Credit Card Pay'
			varNasNbr = NasNbr;
			
			//
			string SQL2 = "SELECT a.service_residential FROM app_request a where app_request_nbr  = " + NasNbr + ";";
			DataTable dt2 = DBquery.RunQuery(SQL2);
			string reqSerType = dt2.Rows[0][0].ToString().Trim();
			
			//Report request submit status
			Report.Log(ReportLevel.Success, "Validation", "Request has been successfully submitted");
			Validate.Exists(repo.DomNasHome.ScotiaBRM.StrongTagViewAppraisalNum);
			Report.Log(ReportLevel.Info, "Information", "Request Number is: " + NasNbr);
			Report.Info("Information","Service type is: " + reqSerType);
			Delay.Milliseconds(100);
			
			
			/*/
			//Report request submit status for 'Credit card to Pay'
			Report.Log(ReportLevel.Success, "Validation", "Request has been successful submit");
			Report.Log(ReportLevel.Info, "Validation", "Request Number is: " + NasNbr);
			Validate.Exists(repo.DomNasHome.ScotiaBRM.YourAssignedNASNumberIs);
			Delay.Milliseconds(100);
			/*/
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
			
		}
	}
}
