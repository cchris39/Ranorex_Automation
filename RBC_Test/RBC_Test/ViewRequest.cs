/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 19/04/2016
 * Time: 10:02 AM
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

namespace RBC_Test
{
	/// <summary>
	/// Description of ViewRequest.
	/// </summary>
	[TestModule("168FDB93-5D90-4DEC-8294-C835BEA393CE", ModuleType.UserCode, 1)]
	public class ViewRequest : ITestModule
	{
		public static RBC_TestRepository repo = RBC_TestRepository.Instance;
		
		static ViewRequest instance = new ViewRequest();
		
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		/// 
		
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("3204ABE6-619F-4AC2-8028-AB9C594C15A7")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varDate = "";
		[TestVariable("37E8A0FB-3CBC-477D-B56C-4136941B8ED6")]
		public string varDate
		{
			get { return _varDate; }
			set { _varDate = value; }
		}
		
		string _varAppName = "";
		[TestVariable("E9EA4347-A724-414E-951F-AFA7B9ED6D51")]
		public string varAppName
		{
			get { return _varAppName; }
			set { _varAppName = value; }
		}
		
		string _varCasperNbr = "";
		[TestVariable("23F6577D-B217-4F05-A405-4CA6BF161872")]
		public string varCasperNbr
		{
			get { return _varCasperNbr; }
			set { _varCasperNbr = value; }
		}
		
		string _varPoCode = "";
		[TestVariable("4F295CB9-02B7-44BA-81B4-2A4E574E2328")]
		public string varPoCode
		{
			get { return _varPoCode; }
			set { _varPoCode = value; }
		}
		
		string _varAddress = "";
		[TestVariable("56610B60-8C02-4BB5-BE72-43E2C190B996")]
		public string varAddress
		{
			get { return _varAddress; }
			set { _varAddress = value; }
		}
		
		string _varStatus = "";
		[TestVariable("B5AEAFCB-FC1E-47C4-AE52-2B8438018C34")]
		public string varStatus
		{
			get { return _varStatus; }
			set { _varStatus = value; }
		}
		
		
		string _varUid = "";
		[TestVariable("87821219-868A-4D1F-ACBD-1F9316C7228F")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("66F1C195-4425-4804-BDF7-7CD8C26CB3AF")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varUrl = "";
		[TestVariable("20795F62-7287-40BB-B7EC-001160696CFA")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		
		#endregion
		
		private string uid = "testrbc@nas.com";
    	private string pwd = "xxxxxxx";
    	private string url = "https://uattest.nas.com/NAS/rbcpvpnas";
    	private string nasNbr, status, address, regDate, appName, contactName, contactNbr, spDir, price, loanType, mortg, branch, property, serviceType, casperNbr;
    	private string adss, city, prov, poCode, result;
    	
    	
		
		public ViewRequest()
		{
			// Do not delete - a parameterless constructor is required!
		}

		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		/// 
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
		
		public string[] viewRequestInfo()
		{
			string nasNbr = repo.RBC_Portal.ViewRequest.ViewNasNbrTag.InnerText.Trim();
			string status = repo.RBC_Portal.ViewRequest.RequestStatus.InnerText.Trim();
			string address = repo.RBC_Portal.ViewRequest.PropertyAddress.InnerText.Trim();
			string reqDate = repo.RBC_Portal.ViewRequest.RequestDate.InnerText.Trim();
			string appName = repo.RBC_Portal.ViewRequest.ApplicantName.InnerText.Trim();
			string contactName = repo.RBC_Portal.ViewRequest.ConatctName.InnerText.Trim();
			string contactNbr = repo.RBC_Portal.ViewRequest.ContactNbr_1st.InnerText.Trim();
			string spDir = repo.RBC_Portal.ViewRequest.SpDirectionText.InnerText.Trim();
			string price = repo.RBC_Portal.ViewRequest.Price.InnerText.Trim();
			string loanType = repo.RBC_Portal.ViewRequest.LoanType.InnerText.Trim();
			string mortg = repo.RBC_Portal.ViewRequest.MortgageAmount.InnerText.Trim();
			string branch = repo.RBC_Portal.ViewRequest.TransitNbr.InnerText.Trim();
			string property = repo.RBC_Portal.ViewRequest.PropertyType.InnerText.Trim();
			string serviceType = repo.RBC_Portal.ViewRequest.ServiceType.InnerText.Trim();
			string casperNbr = repo.RBC_Portal.ViewRequest.CASPER_Nbr.InnerText.Trim();
			
			string[] appAddress = address.Split(',');
			string adss = appAddress[0].Trim();
			string city = appAddress[1].Trim();
			string prov = appAddress[2].Trim();
			string poCode = appAddress[3].Trim();
			reqDate = reqDate.Substring(0,10);
			
			string[] viewResult = new string[] {nasNbr, status, adss, reqDate, city, prov, poCode, appName, contactName, contactNbr, spDir, price, loanType, mortg, branch, property, serviceType, casperNbr};
			
			return viewResult;
		}
		
	
		
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			//RBC user log in
			Login userLogin = new Login();
			userLogin.login(varUid, varPwd, varUrl);
			
			Delay.Milliseconds(100);
			
			//Search request to get to the view request page
			repo.RBC_Portal.SearchFilter.Click();
			repo.RBC_Portal.SearchRequest.NasReqNum.PressKeys(varNasNbr);
			repo.RBC_Portal.SearchRequest.ViewGroupReq.Click();
			repo.RBC_Portal.SearchRequest.SearchSubmitBtn.Click();
			repo.RBC_Portal.SearchRequest.SelectRadioRow.Click();
			repo.RBC_Portal.SearchRequest.ViewRequest.Click();
			
			string address = repo.RBC_Portal.ViewRequest.PropertyAddress.InnerText.Trim();
			
			string SQL = "SELECT a.app_request_nbr, a.`status`, a.`app_address`,  a.`req_date`, a.`app_city`, a.`app_province`, a.`app_postal_code`, a.`app_name`, a.`app_contact_name`," +
						"a.`app_contact_tel`, a.`app_special_dir`, a.`price`, a.`price_type`, a.`mortgage_amt`, a.`branch_no`, a.`property_type`, a.`service_residential`, b.value " +
						"FROM app_request a	join app_request_ext b on a.app_request_nbr = b.app_request_nbr where b.app_request_ext_type_id = 5 and a.app_request_nbr = " + varNasNbr + ";";
			
			
			DataTable dt = QueryDB.RunQuery(SQL);
			
			
			string[] reqInfo = new string[18];
			
			//Construct a string array contain the information from DB
			for (int i = 0; i <= 17; i++)
			{
				reqInfo[i] = dt.Rows[0][i].ToString().Trim();
				
				//Format the string for reqDate, Price, Mortgage
				if (i==3) 
				{
					reqInfo[i] = string.Format("{0:u}", dt.Rows[0][i]).Substring(0,10);
					varDate = reqInfo[3];
				}else if (i==11 || i==13)
				{	reqInfo[i] = reqInfo[i] + ".00";}
				else if (reqInfo[i] == "Purchase Price")
				{ reqInfo[i] = "Purchase";}
				else if (reqInfo[i] == "Refinance Amount")
				{ reqInfo[i] = "Refinance";}
			}
			
			//Assign value to module varilables for searching request
			varStatus = reqInfo[1];
			varAddress = reqInfo[2];
			varDate = reqInfo[3];
			varPoCode = reqInfo[6];
			varAppName = reqInfo[7];
			varCasperNbr = reqInfo[17];
			//ViewNasNbr = varNasNbr;
			
		
			//Get request information from view request page
			string[] viewResult = viewRequestInfo();
			
			//Compare with both information
			bool assert1 = ArraysAreEqual(viewResult, reqInfo);
			
			
			//Format Empty String output to 'Empty' and display in Report
			for (int i = 0; i <= 17; i++)
			{
				string check = CheckEmpty(viewResult[i]);
				if(check == "Empty" )
				{viewResult[i] = check;}
			}
			
			//Report result
			if(assert1)
			{
				Report.Log(ReportLevel.Success, "Success", "Order" + nasNbr + " information verified, request information matched." );
				Report.Log(ReportLevel.Info, "Information", "Nas Request number: " + PutIntoQuotes(viewResult[0]));
				Report.Log(ReportLevel.Info, "Information", "Request status: " + PutIntoQuotes(viewResult[1]));
				Report.Log(ReportLevel.Info, "Information", "Property address: " + PutIntoQuotes(viewResult[2]));
				Report.Log(ReportLevel.Info, "Information", "Request date: " + PutIntoQuotes(viewResult[3]));
				Report.Log(ReportLevel.Info, "Information", "Property city: " + PutIntoQuotes(viewResult[4]));
				Report.Log(ReportLevel.Info, "Information", "Property province: " + PutIntoQuotes(viewResult[5]));
				Report.Log(ReportLevel.Info, "Information", "Property postal code: " + PutIntoQuotes(viewResult[6]));
				Report.Log(ReportLevel.Info, "Information", "Applicant name: " + PutIntoQuotes(viewResult[7]));
				Report.Log(ReportLevel.Info, "Information", "Contact name: " + PutIntoQuotes(viewResult[8]) );
				Report.Log(ReportLevel.Info, "Information", "Contact phone number: " + PutIntoQuotes(viewResult[9]));
				Report.Log(ReportLevel.Info, "Information", "Special direction: " + PutIntoQuotes(viewResult[10]));
				Report.Log(ReportLevel.Info, "Information", "Purchase value: " + PutIntoQuotes(viewResult[11]));
				Report.Log(ReportLevel.Info, "Information", "Request Loan type: " + PutIntoQuotes(viewResult[12]));
				Report.Log(ReportLevel.Info, "Information", "Mortgage amount: " + PutIntoQuotes(viewResult[13]));
				Report.Log(ReportLevel.Info, "Information", "Brnch transit number: " + PutIntoQuotes(viewResult[14]));
				Report.Log(ReportLevel.Info, "Information", "Property type: " + PutIntoQuotes(viewResult[15]));
				Report.Log(ReportLevel.Info, "Information", "Request service type is: " + PutIntoQuotes(viewResult[16]));
				Report.Log(ReportLevel.Info, "Information", "Request CASPER/Linx number is: " + PutIntoQuotes(viewResult[17]));
			}
			else
			{Report.Log(ReportLevel.Failure, "Fail", "Order" + nasNbr + " information verified, request information not match." );}
			
			
			//Close Browser
			//Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
		}
	}
}
