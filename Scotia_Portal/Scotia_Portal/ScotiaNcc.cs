/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-05-16
 * Time: 4:20 PM
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

namespace Scotia_Portal
{
	/// <summary>
	/// Description of ScotiaNcc.
	/// </summary>
	[TestModule("762B8EC2-B2B1-473A-A6B6-C1B415E0052B", ModuleType.UserCode, 1)]
	public class ScotiaNcc : ITestModule
	{
		
		public static Scotia_PortalRepository repo = Scotia_PortalRepository.Instance;
		
		static ScotiaNcc instance = new ScotiaNcc();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ScotiaNcc()
		{
			// Do not delete - a parameterless constructor is required!
		}
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("494CACB3-3E7A-4C48-86A5-607783AC06B0")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varPoCode = "";
		[TestVariable("376A5E8E-D7E9-431C-B53D-FA32F37574F7")]
		public string varPoCode
		{
			get { return _varPoCode; }
			set { _varPoCode = value; }
		}
		
		string _varService = "";
		[TestVariable("3A03A2A7-8C8F-40E7-BA2F-849489AA55A1")]
		public string varService
		{
			get { return _varService; }
			set { _varService = value; }
		}
		
		string _varAppraisal = "";
		[TestVariable("65F0546E-7006-4763-9FE3-BD9353587481")]
		public string varAppraisal
		{
			get { return _varAppraisal; }
			set { _varAppraisal = value; }
		}
		
		string _varSerValue = "";
		[TestVariable("36A63FC3-5B23-4258-A596-D343E7F057B4")]
		public string varSerValue
		{
			get { return _varSerValue; }
			set { _varSerValue = value; }
		}
		
		string _varSaOption = "";
		[TestVariable("67DF338C-4A78-436B-973A-9D1F3DDCDF85")]
		public string varSaOption
		{
			get { return _varSaOption; }
			set { _varSaOption = value; }
		}
		
		string _varLoan = "";
		[TestVariable("096B37BA-BEBD-47BC-A421-61BCF6D21926")]
		public string varLoan
		{
			get { return _varLoan; }
			set { _varLoan = value; }
		}
		
		string _varProperty = "";
		[TestVariable("2BB92AA2-2F6F-4208-9F1A-15907B457EE2")]
		public string varProperty
		{
			get { return _varProperty; }
			set { _varProperty = value; }
		}
		
		string _varUid = "";
		[TestVariable("5673D36A-9837-4A7F-BB54-82DBCF650B8A")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("8FFA4FEB-B972-4D62-8FFC-4018C19FABA9")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		
		string _varIVP_ServiceType = "";
		[TestVariable("B252189B-5080-4506-9D71-DC91B900C161")]
		public string varIVP_ServiceType
		{
			get { return _varIVP_ServiceType; }
			set { _varIVP_ServiceType = value; }
		}
		
		#endregion
		
		private const string addiEmail = "adminuattest.nas.com";
		private const string appName = "Ncc Applicant";
		private const string contact = "Ncc Contact";
		private const string spText = "Scotia NCC test.";
		private const string strName = "Stassen";
		private const string strType = "Court"; 
		private string price, oriNbr, orderSerType, reqServiceType;
		private string[] arrVals;
		
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		
		public void selectAppraisal(string appraisalType)
		{
			switch (appraisalType)
				{
					case "PropertyGauge":
					repo.DomScotia.RequestService.PropertyGauge.Click();
       				    break;
       				case "Residential New Request":
       				    repo.DomScotia.RequestService.ResidentialNewRequest.Click();
       				    break;
       				case "Residential Update":
       				    repo.DomScotia.RequestService.ResidentialUpdate.Click();
       				    break;
       				case "Small Business":
       				    repo.DomScotia.RequestService.SmallBusiness.Click();
       				    break;
       				default:
        			    MessageBox.Show("Please specify request appraisal type.");
        			break;    
				}
		
		}  //End of selectAppraisal
		
		public void selectService(string serviceType)
		{
			Delay.Milliseconds(100);
			if(serviceType != "PropertyGauge")
			{
				Ranorex.InputTag selSertype = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[#'$']".Replace("$", serviceType).Trim();
				selSertype.Click();
				Delay.Milliseconds(100);
			}
		}   //End of Select Request
		
		
		public void requestInfo (string appraisalType, string serviceType, string postal, string strNbr, string strName, string strType, string originalNbr,
		                         string loan, string priceValue, string property, string serValue, string saOption, string refNbr, string appName, 
		                         string contact, string tel1, string tel2, string tel3, string spTxt, string email)
		{
				if(appraisalType != "PropertyGauge")        //[Appraisal Service Type]
				{
					//Select service type [Residential New Request]
					Ranorex.InputTag selService = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[#'$']".Replace("$", serviceType).Trim();
					Delay.Milliseconds(100);
					selService.Click();
				}
				
					if(appraisalType != "Residential Update")
					{	
					//Property Address information is required  [Property Address]
					//Property address information
					repo.DomScotia.RequestService.PropertyAddress.AppPostalCode.PressKeys(postal);
					repo.DomScotia.RequestService.PropertyAddress.AppPostalCode.PressKeys("{Tab}");
					Delay.Milliseconds(100);
				
					repo.DomScotia.RequestService.PropertyAddress.AppStreetNum.PressKeys(strNbr);
					repo.DomScotia.RequestService.PropertyAddress.AppStreetName.PressKeys(strName);
					repo.DomScotia.RequestService.PropertyAddress.AppStreetType.PressKeys(strType);
					Delay.Milliseconds(100);
					}else
					//If Service Tyep is "Residential Update"
					//if (appraisalType == "Residential Update" )    //|| serviceType == "ProgressAdvance"
					{
						Delay.Milliseconds(100);
						//Message pop op for these two service types
						if(serviceType == "AppraisalupdateAmendmentofExistingReport" || serviceType == "AddRentalScheduleAtoExistingReport")
						{
							Delay.Milliseconds(100);
							repo.MessageFromWebpage.ButtonOK.Click();
						}
						Delay.Milliseconds(100);
						//Get original number for Residential Update
						repo.DomScotia.RequestService.Ncc_RequestService.OriginalReqNbr.PressKeys(originalNbr);
						Delay.Milliseconds(200);
						repo.DomScotia.RequestService.Ncc_RequestService.OriNbrSearchBtn.PerformClick();
						Delay.Milliseconds(100);
					}
					
					
				//	[Appraisal Purpose]
				if ( appraisalType != "PropertyGauge" && appraisalType !="Residential Update" && serviceType != "ConditionReport" && serviceType != "ProgressAdvance")
				{
					//Select Appraisal Purpose
					Ranorex.InputTag selAppPurpose = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[#'$']".Replace("$", loan).Trim();
					selAppPurpose.Click();
					Delay.Milliseconds(100);
					
					//Input Estimate Value
					repo.DomScotia.RequestService.Ncc_RequestInfo.EstimatedValue.PressKeys(priceValue);
				}
					// Known Environment in Appraisal Purpose
					repo.DomScotia.RequestService.Ncc_RequestInfo.KnownEnvironmentalIssue_No.Click();
					Delay.Milliseconds(100);
					
					
				//Property Type options   [Property Type options]
				if (appraisalType == "Residential New Request")
				{
					Ranorex.InputTag selPropertyType = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[#'$']".Replace("$", property).Trim();
					selPropertyType.Click();
					Delay.Milliseconds(100);
				
					//Service Type options
					if(serviceType == "Full-Service")
					{
						//Choose 'Service Type' options
						Ranorex.InputTag selSerValue = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[#'$']".Replace("$",serValue).Trim();
						selSerValue.Click();
						Delay.Milliseconds(100);
						
						//Pop op message for As completed and Both Values
						if(serValue == "AsCompleteValueOnlyProspective" || serValue == "BothAsisandAsCompleteProspectiveValues" )
						{
							repo.MessageFromWebpage.ButtonOK.Click();
						}
					}else if(serviceType != "ProgressAdvance" )
					{
							//Default to "As is "
							repo.DomScotia.RequestService.Ncc_RequestInfo.AsisValueOnly.Click();
							Delay.Milliseconds(100);
					}
					
						//Schedule A and Rush Options
						Ranorex.InputTag sa = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[#'scheduleappraisal$']".Replace("$", saOption).Trim();
						sa.Click();
						Delay.Milliseconds(100);
						//Rush options
						repo.DomScotia.RequestService.Ncc_RequestInfo.RushserviceNo.PerformClick();
						Delay.Milliseconds(100);
				}  //End of Property type option
				
				
					// [Applicant and contact information]
					if (appraisalType != "Residential Update" )
					{	
						//Loan and Contact Information
						repo.DomScotia.RequestService.LoanAndContactInformation.AppRefNbr.PressKeys(refNbr);
						repo.DomScotia.RequestService.LoanAndContactInformation.AppName.PressKeys(appName);
						if(serviceType != "PropertyGauge" && serviceType != "Drive-By" && serviceType != "Desktop")      
						{	//Contact Information
						repo.DomScotia.RequestService.LoanAndContactInformation.Ncc_AppContactName.PressKeys(contact);
						repo.DomScotia.RequestService.LoanAndContactInformation.Tel1.PressKeys(tel1);
						repo.DomScotia.RequestService.LoanAndContactInformation.Tel2.PressKeys(tel2);
						repo.DomScotia.RequestService.LoanAndContactInformation.Tel3.PressKeys(tel3);
						repo.DomScotia.RequestService.LoanAndContactInformation.PreferNumber1.Click();
						
							//Special Instructions and email notifications
							if (serviceType != "ProgressAdvance")
							{
							repo.DomScotia.RequestService.SpecialInstructions.SpText.PressKeys(spTxt);
							repo.DomScotia.RequestService.EmailNotifications.AdditionalEmail.PressKeys(email);
							Delay.Milliseconds(100);
							}
						}
					} //End of Applicant and contact information
		
				//Submit appraisal request
				repo.DomScotia.RequestService.SubmitRequest.PerformClick();
				Delay.Milliseconds(200);
				
				if(appraisalType == "PropertyGauge")
				{
					if(repo.DomScotia.RequestService.Ncc_RequestInfo.CheckAddressInfo.Exists())
					{
					repo.DomScotia.RequestService.Ncc_RequestInfo.CheckAddress.Click();
					Delay.Milliseconds(100);
					
					repo.DomScotia.RequestService.Ncc_RequestInfo.SubmitPropertyGauge.PerformClick();
					Delay.Milliseconds(300);
					}
				}else if(appraisalType != "Residential Update" || serviceType!= "UpgradeDesktopAppraisaltoFull")
				{	//Select Appraiser before final submit
					//Check if there is any appraiser available to select
					if(repo.DomScotia.RequestService.Ncc_RequestService.SelectAppraiserRadioRowInfo.Exists())
					{repo.DomScotia.RequestService.Ncc_RequestService.NccSelectAppraiserSubmitBtn.PerformClick();
					Delay.Milliseconds(300);
					}
				}else
				{ Delay.Milliseconds(100);}
		
		}//End of requestInfo
		
		
		public string[] getOriNbr(string serviceType, string uid)
		{
			string SQL1 = "SELECT a.app_request_nbr FROM app_request a where username = '" + uid + "' and service_residential != 'Property Gauge' and service_residential !='P.A. SMC' and status = 'Completed' order by app_request_nbr desc limit 1;";
			string SQL2 = "SELECT a.app_request_nbr, a.service_residential FROM app_request a where username = '" + uid + "' and service_residential != 'Property Gauge' and service_residential != 'P.A. SMC' and service_residential not like '%Schedule A%' and service_residential != 'Appraisal Update' and service_residential != 'Condition Report' and status !='Completed' order by app_request_nbr desc limit 1;";
			string SQL3 = "SELECT a.app_request_nbr FROM app_request a where username = '" + uid + "' and service_residential REGEXP 'Desktop' and STATUS = 'Completed' order by app_request_nbr desc limit 1;";
			string SQL4 = "SELECT a.app_request_nbr FROM app_request a where username = '" + uid + "' and service_residential REGEXP 'P.A. SMC' and STATUS = 'Completed' order by app_request_nbr desc limit 1;";
			
			string[] val = new string[2];
			if (serviceType == "AppraisalupdateAmendmentofExistingReport")
			{
					DataTable dt1 = QueryDB.RunQuery(SQL1);
					oriNbr = dt1.Rows[0][0].ToString();
					val[0] = oriNbr;
					return val;
			}else if(serviceType == "AddRentalScheduleAtoExistingReport")
			{		
					DataTable dt2 = QueryDB.RunQuery(SQL2);
					oriNbr = dt2.Rows[0][0].ToString();
					orderSerType = dt2.Rows[0][1].ToString();
					val[0] = oriNbr;
					val[1] = orderSerType;
					return val;
			}else if (serviceType == "UpgradeDesktopAppraisaltoFull")
			{	
					DataTable dt3 = QueryDB.RunQuery(SQL3);
					oriNbr = dt3.Rows[0][0].ToString();
					val[0] = oriNbr;
					return val;
			}else if (serviceType == "FinalInspection")
			{	
					DataTable dt4 = QueryDB.RunQuery(SQL4);
					oriNbr = dt4.Rows[0][0].ToString();
					val[0] = oriNbr;
					return val;
			}else
			{ 		val[0] = "";
					return val;
			}	
		
		}  // End of getOriNbr
		
		public string getRequestNumber()
		{
			string nbr = repo.DomScotia.StrongTagAppraisalRequestNumber.InnerText.Trim().Substring(30, 7);
			repo.DomScotia.StrongTagAppraisalRequestNumber.Click();
			Delay.Milliseconds(100);
			return nbr;
			
		}    //End of [reportOrderStatus] function
		
		public string getRequestServiceType()
		{
			repo.DomScotia.SearchResults.ViewAppraisalRequest.Click();
			Delay.Milliseconds(200);
			
			string serType = repo.DomScotia.ViewRequest.ServiceType.InnerText.Trim();
			return serType;
			
		}    //End of [reportOrderStatus] function
		
		public string getPropertyGaugeNumber()
		{
			string nbr = repo.DomScotia.PropertyGaugeConfirmationPage.StrongTagViewAppraisalRequestNumbe.InnerText.Trim().Substring(31, 7).Trim();
			repo.DomScotia.PropertyGaugeConfirmationPage.StrongTagViewAppraisalRequestNumbe.Click();
			Delay.Milliseconds(100);
			return nbr;
			
		}    //End of [reportOrderStatus] function
		
		
		
		//=======================================================================================
		
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			Delay.Milliseconds(200);
			
			var random = new Random();
			string stNbr = random.Next(1, 900).ToString();
			string clientRef = random.Next(123456789, 987654321).ToString();
			string tel1 = random.Next(111, 999).ToString();
			string tel2 = random.Next(111, 999).ToString();
			string tel3 = random.Next(1111,9999).ToString();
			int Price = random.Next(500,900)*1000;
			price = Price.ToString();
			
			Login UsrLogin = new Login();
			UsrLogin.lauchScotia();
			UsrLogin.UserLogin(varUid, varPwd);
						
			//Get original Nas Number based on Appraisal Service Type
			arrVals = getOriNbr (varService, varUid);
			oriNbr = arrVals[0];
			orderSerType = arrVals[1];
						
			//Request Service through linked test data table
			repo.DomScotia.MainMenu.RequestService.Click();
			Delay.Milliseconds(200);
			
			//Select Appraisal Type
			selectAppraisal(varAppraisal);
			Delay.Milliseconds(100);
			
			//Select Service Type
			selectService(varService);
			Delay.Milliseconds(100);
			
			//Fill out request information based on appraisal and service type
			requestInfo(varAppraisal, varService, varPoCode, stNbr, strName, strType, oriNbr, varLoan, price, varProperty, 
			            varSerValue, varSaOption, clientRef, appName, contact, tel1, tel2, tel3, spText, addiEmail);
			Delay.Milliseconds(100);
			
			if(varAppraisal != "PropertyGauge")
			{	//Get Nas number
				varNasNbr = getRequestNumber();
			}else
			{	varNasNbr = getPropertyGaugeNumber(); }
			
			if(varService == "AddRentalScheduleAtoExistingReport")
			{
				//Get Expected Service type
				varIVP_ServiceType = orderSerType + " with Rental Schedule";
			}		
				
				//Get Actual Service type
				reqServiceType = getRequestServiceType();
				
			
			Report.Log(ReportLevel.Info, varNasNbr + " request service type is: " + reqServiceType);
			Validate.AreEqual(reqServiceType, varIVP_ServiceType);
			
			Delay.Milliseconds(100);
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
		}
	}
}
