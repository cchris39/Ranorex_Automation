/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 09/02/2016
 * Time: 9:49 AM
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

namespace NasAdmin
{
	/// <summary>
	/// Description of AdminViewRequest.
	/// </summary>
	[TestModule("3E24620F-5C41-4B45-8BF1-58A66D0A8157", ModuleType.UserCode, 1)]
	public class AdminViewRequest : ITestModule
	{
		/// <summary>
		/// Using the NasAdminRepository repository.
		/// </summary>
		public static NasAdminRepository repo = NasAdminRepository.Instance;
		
		static AdminViewRequest instance = new AdminViewRequest();
		
		#region Varilables
		
		string _varUrl = "";
		[TestVariable("1E239EBC-7629-48BF-AD25-6F31C1A416B1")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varPropertyAddress = "";
		[TestVariable("C56B0F28-CE0D-46FB-92B7-6B70054FE87B")]
		public string varPropertyAddress
		{
			get { return _varPropertyAddress; }
			set { _varPropertyAddress = value; }
		}
		
		
		string _varSerType = "";
		[TestVariable("8FA33E63-6D73-430D-B7B4-012F7175EBD0")]
		public string varSerType
		{
			get { return _varSerType; }
			set { _varSerType = value; }
		}
		
		
		string _varPrice = "";
		[TestVariable("66841644-CE4E-4C9C-87CE-8388EE0E00BF")]
		public string varPrice
		{
			get { return _varPrice; }
			set { _varPrice = value; }
		}
		
		string _varMortgAmount = "";
		[TestVariable("2399041C-D7B5-4D58-8F6E-69EEFABB8AD3")]
		public string varMortgAmount
		{
			get { return _varMortgAmount; }
			set { _varMortgAmount = value; }
		}
		
		string _varPostalCode = "";
		[TestVariable("676C8F13-74F6-4543-BA03-F85C9AEEBDE1")]
		public string varPostalCode
		{
			get { return _varPostalCode; }
			set { _varPostalCode = value; }
		}
		
		string _varCreationDate = "";
		[TestVariable("5F93B261-C892-4D30-AB43-5EF9F4F9197D")]
		public string varCreationDate
		{
			get { return _varCreationDate; }
			set { _varCreationDate = value; }
		}
		
		string _varRefNbr = "";
		[TestVariable("4222356B-B896-4691-941E-C415FDC4326A")]
		public string varRefNbr
		{
			get { return _varRefNbr; }
			set { _varRefNbr = value; }
		}
		
		string _varSpText = "";
		[TestVariable("525481ED-D89C-421C-8913-048CD5116102")]
		public string varSpText
		{
			get { return _varSpText; }
			set { _varSpText = value; }
		}
		
		string _varAppName = "";
		[TestVariable("72978481-7BE1-44A4-96C4-CF4CBB8A68BE")]
		public string varAppName
		{
			get { return _varAppName; }
			set { _varAppName = value; }
		}
		
		string _varContactName = "";
		[TestVariable("E00C5E8C-7BBA-40D6-8873-B07401CEB25F")]
		public string varContactName
		{
			get { return _varContactName; }
			set { _varContactName = value; }
		}
		
		string _varContactTel = "";
		[TestVariable("04B1B82C-A161-4D8B-96AC-7D37D8C0C8DF")]
		public string varContactTel
		{
			get { return _varContactTel; }
			set { _varContactTel = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("E4E26759-7AB8-41DA-8083-4D1A22C9E174")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varBranchNbr = "";
		[TestVariable("0B49C0F5-E613-4AE3-91BE-B2A0EDB82011")]
		public string varBranchNbr
		{
			get { return _varBranchNbr; }
			set { _varBranchNbr = value; }
		}
		
		string _varProvince = "";
		[TestVariable("D1A2C135-1BF3-4062-BB07-46295310DDBC")]
		public string varProvince
		{
			get { return _varProvince; }
			set { _varProvince = value; }
		}
		
		string _varCity = "";
		[TestVariable("011D8EBF-D461-4EAB-8875-229BC2E1902D")]
		public string varCity
		{
			get { return _varCity; }
			set { _varCity = value; }
		}
		
		#endregion
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public AdminViewRequest()
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
			
			
			
			//Query Request information from DB
			string SQL1 = "SELECT a.`appraiser`, a.`app_city`, a.`app_province`, a.`price_type` FROM app_request a where app_request_nbr = " + varNasNbr ;       //varNasNbr
			DataTable dt= DatabaseQuery.RunQuery(SQL1);
			
			string appraiser = dt.Rows[0][0].ToString().Trim();
			string city = dt.Rows[0][1].ToString().Trim();
			string province = dt.Rows[0][2].ToString().Trim();
			string loanType = dt.Rows[0][3].ToString().Trim();
			
			varProvince = province;
			varCity = city;
			
			string SQL2 = "SELECT u.`first_name`, u.`last_name` FROM user_info u where username = " + "'" + appraiser + "';";                //appraiser
			DataTable dt2 = DatabaseQuery.RunQuery(SQL2);
			
			string aprFstName = dt2.Rows[0][0].ToString().Trim();
			string aprLstName = dt2.Rows[0][1].ToString().Trim();
			string appraiserName = aprFstName + aprLstName;        //Add 3 space in between
						
			//Search Nas Number
			repo.Dom_NasHome.NasAdminMenu.SearchFilter.Click();
			repo.Dom_NasHome.MenuDisplay.NasReqNumSearch.PressKeys(varNasNbr);      //varNasNbr
			repo.Dom_NasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(200);
			
			//View Request
			repo.Dom_NasHome.NasAdminFunction.ViewRequest.Click();
			Delay.Milliseconds(200);
			
			//Get Request information from View Request Page
			string aprName = repo.Dom_NasHome.MenuDisplay.ViewRequest.AppraiserName.InnerText.Trim();
			string serviceType = repo.Dom_NasHome.MenuDisplay.ViewRequest.ServiceType.InnerText.Trim();
			string price = repo.Dom_NasHome.MenuDisplay.ViewRequest.EstimatedValue.InnerText.Trim();
			string mortgage = repo.Dom_NasHome.MenuDisplay.ViewRequest.MortgageAmt.InnerText.Trim();
			string priceType = repo.Dom_NasHome.MenuDisplay.ViewRequest.LoanType.InnerText.Trim();
			string status = repo.Dom_NasHome.MenuDisplay.ViewRequest.Status.InnerText.Trim();
			string cDate = repo.Dom_NasHome.MenuDisplay.ViewRequest.ReqDate.InnerText.Trim().Substring(0, 10);
			string spTxt = repo.Dom_NasHome.MenuDisplay.ViewRequest.SpecialDirections.InnerText.Trim();
			string braNbr = repo.Dom_NasHome.MenuDisplay.ViewRequest.BranchNbr.Value.Trim();
			string appName = repo.Dom_NasHome.MenuDisplay.ViewRequest.AppName.InnerText.Trim();
			string conName = repo.Dom_NasHome.MenuDisplay.ViewRequest.ContactName.InnerText.Trim();
			string conTel = repo.Dom_NasHome.MenuDisplay.ViewRequest.ContactPhoneNbr.InnerText.Trim();
			string address = repo.Dom_NasHome.MenuDisplay.ViewRequest.PropertyAddress.InnerText.Trim();
			string reqCity = repo.Dom_NasHome.MenuDisplay.ViewRequest.City.InnerText.Trim();
			string reqProv = repo.Dom_NasHome.MenuDisplay.ViewRequest.Province.TagValue;
			string postalCode = repo.Dom_NasHome.MenuDisplay.ViewRequest.Req_PostalCode.InnerText.Trim();
			string reqNbr = repo.Dom_NasHome.MenuDisplay.ViewRequest.NASRequestNumber.InnerText.Trim().Substring(22,10).Trim();
			string refNbr = repo.Dom_NasHome.MenuDisplay.ViewRequest.ReqRefNbr.InnerText.Trim();
			
			
			//Validate the information against request ordered and DB
			//Report.Log(ReportLevel.Success, "Validation", "Appraiser information is match.");
			//Validate.AreEqual(appraiserName, aprName);  
			
			Report.Log(ReportLevel.Success, "Validation", "Nas request number is match.");
			Validate.AreEqual(varNasNbr, reqNbr);     //varNasNbr
			
			Report.Info("Request service type is: " + serviceType);
			Report.Info("Request status is : " + status);
			
			Report.Log(ReportLevel.Success, "Validation", "Price information is match.");
			Validate.AreEqual(varPrice, price);   
			
			Report.Log(ReportLevel.Success, "Validation", "Loan type is match.");
			Validate.AreEqual(loanType, priceType);  
			
			Report.Log(ReportLevel.Success, "Validation", "Mortgage amount is match.");
			Validate.AreEqual(varMortgAmount, mortgage);  
			
			Report.Log(ReportLevel.Success, "Validation", "Property address is match.");
			Validate.AreEqual(varPropertyAddress, address);  
			
			Report.Log(ReportLevel.Success, "Validation", "Property city is match.");
			Validate.AreEqual(city, reqCity);  
			
			Report.Log(ReportLevel.Success, "Validation", "Property province is match.");
			Validate.AreEqual(province, reqProv);  
			
			Report.Log(ReportLevel.Success, "Validation", "Postal Code is match.");
			Validate.AreEqual(varPostalCode, postalCode);  
			
			Report.Log(ReportLevel.Success, "Validation", "Request creation date is match.");
			Validate.AreEqual(varCreationDate, cDate);  
			
			Report.Log(ReportLevel.Success, "Validation", "Request client branch number is match.");
			Validate.AreEqual(varBranchNbr, braNbr);  
		
			Report.Log(ReportLevel.Success, "Validation", "Request client reference number is match.");
			Validate.AreEqual(varRefNbr, refNbr);  
		
			Report.Log(ReportLevel.Success, "Validation", "Request special direction is match.");
			Validate.AreEqual(varSpText, spTxt);  
			
			Report.Log(ReportLevel.Success, "Validation", "Request applicant name is match.");
			Validate.AreEqual(varAppName, appName);  
			
			Report.Log(ReportLevel.Success, "Validation", "Request contact name is match.");
			Validate.AreEqual(varContactName, conName);  
			
			Report.Log(ReportLevel.Success, "Validation", "Request contact phone number is match.");
			Validate.AreEqual(varContactTel, conTel);  
			
			Delay.Milliseconds(300);
			
		}
	}
}
