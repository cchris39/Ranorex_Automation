/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-05-10
 * Time: 9:54 AM
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
	/// Description of ViewRequest.
	/// </summary>
	[TestModule("58E9DC5D-AFB7-4E33-9BF2-BBD7250D7481", ModuleType.UserCode, 1)]
	public class ViewRequest : ITestModule
	{
		public static Scotia_PortalRepository repo = Scotia_PortalRepository.Instance;
		
			static ViewRequest instance = new ViewRequest();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ViewRequest()
		{
			// Do not delete - a parameterless constructor is required!
		}
		
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("19557692-E524-4BE2-9E6D-802050B4B559")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varDate = "";
		[TestVariable("8826F58D-FFD1-4060-BB40-4D8CF3176734")]
		public string varDate
		{
			get { return _varDate; }
			set { _varDate = value; }
		}
		
		string _varStatus = "";
		[TestVariable("61B3DCDF-1FBC-411D-93D9-418EC6AD1846")]
		public string varStatus
		{
			get { return _varStatus; }
			set { _varStatus = value; }
		}
		
		string _varAppName = "";
		[TestVariable("0C8522F2-C9E3-4B82-8E28-337A1674BA3D")]
		public string varAppName
		{
			get { return _varAppName; }
			set { _varAppName = value; }
		}
		
		string _varRefNbr = "";
		[TestVariable("B342118D-896D-4F26-9339-78DE166A8373")]
		public string varRefNbr
		{
			get { return _varRefNbr; }
			set { _varRefNbr = value; }
		}
		
		string _varPostalCode = "";
		[TestVariable("9E3A9154-8225-4557-B1F1-9543CD96A836")]
		public string varPostalCode
		{
			get { return _varPostalCode; }
			set { _varPostalCode = value; }
		}
		
		string _varAddress = "";
		[TestVariable("7D9D51F8-D460-4A0E-9CDA-2F76BE4A2E30")]
		public string varAddress
		{
			get { return _varAddress; }
			set { _varAddress = value; }
		}
		
		string _varCity = "";
		[TestVariable("067969E5-BB1F-4441-AD73-7115F133F5F3")]
		public string varCity
		{
			get { return _varCity; }
			set { _varCity = value; }
		}
		
		string _varProv = "";
		[TestVariable("6E72B0A6-1D44-4F40-8C72-462289053C87")]
		public string varProv
		{
			get { return _varProv; }
			set { _varProv = value; }
		}
		
		string _varUid = "";
		[TestVariable("D01ACC65-F14C-4D64-8990-EA2C9AC0A835")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("A6887004-C295-4222-930B-7781BD9AD7FB")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		#endregion

		private string nasNbr, status,reqdate,address, city, province, poCode, appName, contactName, tel,spDirTxt, serType, loanType,priceValue, morgAmount, refNbr, branchNbr;
		private string property, serviceoption, addiEmail, saOption, rushOption, other1, other2, other3, other4, other5, other6;
		
		public string [] viewRequestInfo()
		{
			string nasNbr = repo.DomScotia.ViewRequest.NASRequestNumber.InnerText.Substring(22,10).Trim();
			
			string status = repo.DomScotia.ViewRequest.Status.InnerText.Trim();
			string reqDate = repo.DomScotia.ViewRequest.DateRequested.InnerText.Trim();
			string address = repo.DomScotia.ViewRequest.PropertyAddress.InnerText.Trim();
			string city = repo.DomScotia.ViewRequest.City.InnerText.Trim();
			string province = repo.DomScotia.ViewRequest.Province.InnerText.Trim();
			string poCode = repo.DomScotia.ViewRequest.PostalCode.InnerText.Trim();
			string appName = repo.DomScotia.ViewRequest.ApplicantName.InnerText.Trim();
			string contactName = repo.DomScotia.ViewRequest.ContactName.InnerText.Trim();
			string tel = repo.DomScotia.ViewRequest.TelephoneHomeNo.InnerText.Substring(0, 12).Trim();
			string spDirTxt = repo.DomScotia.ViewRequest.SpDirections.InnerText.Trim();
			string serType = repo.DomScotia.ViewRequest.ServiceType.InnerText.Trim();
			string loanType = repo.DomScotia.ViewRequest.LoanType.InnerText.Trim();
			string priceValue = repo.DomScotia.ViewRequest.EstimateValue.InnerText.Trim();
			string morgAmount = repo.DomScotia.ViewRequest.MortgageAmount.InnerText.Trim();
			string refNbr = repo.DomScotia.ViewRequest.ApplicationNbr.InnerText.Trim();
			string branchNbr = repo.DomScotia.ViewRequest.BranchNo.InnerText.Trim();
			
			string property = repo.DomScotia.ViewRequest.PropertyType.InnerText.Trim();
			string serviceOption = repo.DomScotia.ViewRequest.ServiceTypeOptions.InnerText.Trim();
			string addiEmail = repo.DomScotia.ViewRequest.AdditionalEmail.InnerText.Trim();
			string saOption = repo.DomScotia.ViewRequest.SAoption.InnerText.Trim();
			string rushOption = repo.DomScotia.ViewRequest.RushOption.InnerText.Trim();
			
			string other1 = repo.DomScotia.ViewRequest.Other_LeaseLand.InnerText.Trim();
			string other2 = repo.DomScotia.ViewRequest.Other_10Acres.InnerText.Trim();
			string other3 = repo.DomScotia.ViewRequest.Other_KnownEnvironmentIssue.InnerText.Trim();
			string other4 = repo.DomScotia.ViewRequest.Other_TenantOccupied.InnerText.Trim();
			string other5 = repo.DomScotia.ViewRequest.Other_SpecialtyProgam.InnerText.Trim();
			string other6 = repo.DomScotia.ViewRequest.Other_OwenedMoreThan12mths.InnerText.Trim();
			
			string[] viewResult = new string [] {nasNbr, status, reqDate, address, city, province, poCode, appName,
				contactName, tel,spDirTxt, serType, loanType,priceValue, morgAmount, refNbr, branchNbr, property, 
				serviceOption, addiEmail, saOption, rushOption, other1, other2, other3, other4, other5, other6};
			
			return viewResult;
		
		}
		
		public static String CheckEmpty(string s)
    	{
    		if (String.IsNullOrEmpty(s)) 
        	return "Empty";
    		else
        	return String.Format("(\"{0}\") is not empty", s);
    	}
		
		static bool ArraysAreEqual(string[] x, string[] y) 
		    {
            if (x.Length != y.Length) {
                return false;
            }

            for (int index = 0; index < x.Length; index++)
            {
                if (x[index] != y[index]) {
                    return false;
                }
            }

            return true;
        	}
		 
		 public static string PutIntoQuotes(string value)
        {
            return "\"" + value + "\"";
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
			
			Login UsrLogin = new Login();
			UsrLogin.lauchScotia();
			UsrLogin.UserLogin(varUid, varPwd);
			Delay.Milliseconds(100);
			
			//Search Request and go to View Request page
			repo.DomScotia.MainMenu.SearchFilter.Click();
			repo.DomScotia.SearchFilter.NasReqNum.PressKeys(varNasNbr);
			repo.DomScotia.SearchFilter.ViewGroupReq.Click();
			repo.DomScotia.SearchFilter.SearchSubmit.Click();
			repo.DomScotia.SearchResults.ViewAppraisalRequest.Click();
			Delay.Milliseconds(100);
			
			
			
			string SQL = "SELECT a.app_request_nbr , a.status, a.req_date, a.app_address, a.app_city, a.app_province, a.app_postal_code, a.app_name, a.app_contact_name, a.app_contact_tel, a.app_special_dir, a.service_residential, a.price_type, a.price, a.mortgage_amt, a.app_ref_nbr, a.branch_no, b.property_type_option, b.service_type_option,b.additional_email, b.schedule_a_option, b.rush_service, b.propety_on_leased_land, b.land_greater_than_10_acres, b.environmental_building_condition_issue, b.100_pecent_rented_tenant_occupied, b.specialty_application_program, b.property_owned_more_12_months FROM app_request a join app_request_scotia b on a.app_request_nbr = b.app_request_nbr where a.app_request_nbr  = " + varNasNbr;
			
			DataTable dt = QueryDB.RunQuery(SQL);
			
			
			string[] reqInfo = new string[28];
			
			//Construct a string array contain the information from DB
			for (int i = 0; i <= 27; i++)
			{
				reqInfo[i] = dt.Rows[0][i].ToString().Trim();
				
				//Format the string for Price, Mortgage, Date, Service Type
				if (i==2) 
				{
					reqInfo[i] = string.Format("{0:u}", dt.Rows[0][i]).Substring(0,19);
					varDate = reqInfo[2];
				}else if (i==13 | i==14)
				{
					reqInfo[i] = "$" + reqInfo[i] + ".0";
				}else if (i==11)
				{
					reqInfo[i] = reqInfo[i].Replace("Schedule A", "Rental Schedule (A)");
				}
			} //End of For loop
			
				
				//Get request information from view request page
				string[] viewResult = viewRequestInfo();
				
				//Compare with both information
				bool assert1 = ArraysAreEqual(viewResult, reqInfo);
			
			
				//Format Empty String output to 'Empty' and display in Report
				for (int i = 0; i <= 27; i++)
				{
					string check = CheckEmpty(viewResult[i]);
					if(check == "Empty" )
					{viewResult[i] = check;}
				}
				
				//Assign value using in Searching Request Module
				varDate = reqInfo[2];
				varAddress = reqInfo[3];
				varProv = reqInfo[4];
				varCity= reqInfo[5];
				varPostalCode = reqInfo[6];
				varAppName = reqInfo[7];
				varRefNbr = reqInfo[15];
				
				
				//Report Compare result
			if(assert1)
			{
				Report.Log(ReportLevel.Success, "Success", "Order" + nasNbr + " information verified, request information matched." );
				Report.Log(ReportLevel.Info, "Information", "Nas Request number: " + PutIntoQuotes(viewResult[0]));
				Report.Log(ReportLevel.Info, "Information", "Request status: " + PutIntoQuotes(viewResult[1]));
				Report.Log(ReportLevel.Info, "Information", "Request date: " + PutIntoQuotes(viewResult[2]));
				Report.Log(ReportLevel.Info, "Information", "Property address: " + PutIntoQuotes(viewResult[3]));
				Report.Log(ReportLevel.Info, "Information", "Property city: " + PutIntoQuotes(viewResult[4]));
				Report.Log(ReportLevel.Info, "Information", "Property province: " + PutIntoQuotes(viewResult[5]));
				Report.Log(ReportLevel.Info, "Information", "Property postal code: " + PutIntoQuotes(viewResult[6]));
				Report.Log(ReportLevel.Info, "Information", "Applicant name: " + PutIntoQuotes(viewResult[7]));
				Report.Log(ReportLevel.Info, "Information", "Contact name: " + PutIntoQuotes(viewResult[8]) );
				Report.Log(ReportLevel.Info, "Information", "Contact phone number: " + PutIntoQuotes(viewResult[9]));
				Report.Log(ReportLevel.Info, "Information", "Special direction: " + PutIntoQuotes(viewResult[10]));
				Report.Log(ReportLevel.Info, "Information", "Service type: " + PutIntoQuotes(viewResult[11]));
				Report.Log(ReportLevel.Info, "Information", "Loan type: " + PutIntoQuotes(viewResult[12]));
				Report.Log(ReportLevel.Info, "Information", "Estimate value: " + PutIntoQuotes(viewResult[13]));
				Report.Log(ReportLevel.Info, "Information", "Mortgage amount: " + PutIntoQuotes(viewResult[14]));
				Report.Log(ReportLevel.Info, "Information", "Application number is: " + PutIntoQuotes(viewResult[15]));
				Report.Log(ReportLevel.Info, "Information", "Brnch transit number: " + PutIntoQuotes(viewResult[16]));
				Report.Log(ReportLevel.Info, "Information", "Property type: " + PutIntoQuotes(viewResult[17]));
				Report.Log(ReportLevel.Info, "Information", "Service type option is: " + PutIntoQuotes(viewResult[18]));
				Report.Log(ReportLevel.Info, "Information", "Additional Email is: " + PutIntoQuotes(viewResult[19]));
				Report.Log(ReportLevel.Info, "Information", "Schedule A option is: " + PutIntoQuotes(viewResult[20]));
				Report.Log(ReportLevel.Info, "Information", "Rush option is: " + PutIntoQuotes(viewResult[21]));
				Report.Log(ReportLevel.Info, "Information", "Property on leased land is: " + PutIntoQuotes(viewResult[22]));
				Report.Log(ReportLevel.Info, "Information", "Land grteater than 10 Acres is: " + PutIntoQuotes(viewResult[23]));
				Report.Log(ReportLevel.Info, "Information", "Known environmental condition issue is: " + PutIntoQuotes(viewResult[24]));
				Report.Log(ReportLevel.Info, "Information", "100% occupied by tenant is: " + PutIntoQuotes(viewResult[25]));
				Report.Log(ReportLevel.Info, "Information", "Mortgage speciality program is: " + PutIntoQuotes(viewResult[26]));
				Report.Log(ReportLevel.Info, "Information", "Property owned for more than 12 months is: " + PutIntoQuotes(viewResult[27]));
				
			}
			else
			{Report.Log(ReportLevel.Failure, "Fail", "Order" + nasNbr + " information verified, request information not match." );}
			
				//Close Browser
				//Host.Local.KillBrowser("IE");
				Delay.Milliseconds(200);
				
			}
		}
	}

