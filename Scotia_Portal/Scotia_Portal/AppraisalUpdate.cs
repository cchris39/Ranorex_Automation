/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-05-16
 * Time: 11:33 AM
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
	/// Description of AppraisalUpdate.
	/// </summary>
	[TestModule("9828C79B-7F9F-417A-843D-743BCDB8B382", ModuleType.UserCode, 1)]
	public class AppraisalUpdate : ITestModule
	{
		
		public static Scotia_PortalRepository repo = Scotia_PortalRepository.Instance;
		
		static AppraisalUpdate instance = new AppraisalUpdate();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public AppraisalUpdate()
		{
			// Do not delete - a parameterless constructor is required!
		}
		
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("36B9B712-87BD-488C-880E-826228CE873F")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varUpdateType = "";
		[TestVariable("D6D5C5D3-B036-4C43-9463-9CAA6C677B45")]
		public string varUpdateType
		{
			get { return _varUpdateType; }
			set { _varUpdateType = value; }
		}
		
		string _varUid = "";
		[TestVariable("DE2F06C7-030C-4D87-A1E6-D8C60F199F6C")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("BBCCC122-7E4D-4408-B56D-BDA4782E789C")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varIVP_ServiceType = "";
		[TestVariable("C02A7D4F-48DA-4772-85F9-D1AC8446DF22")]
		public string varIVP_ServiceType
		{
			get { return _varIVP_ServiceType; }
			set { _varIVP_ServiceType = value; }
		}
		
		#endregion
		
		private const string addiEmail = "testadmintest@nas.com";
		private const string appName = "First Update";
		private const string contact = "Second Update";
		private const string spText = "Scotia Residential Update test.";
		private const string strName = "South Star";
		private const string strType = "Road"; 
		private const string contactName = "Hello World";
		
		private string service, oriNbr;
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		/// 
		public void selectUpdate(string updateType, string oriNbr)
		{
			switch (updateType)
			{
				case "Appraisal Update":
					repo.DomScotia.RequestService.AppraisalUpdateRequestType.AppraisalupdateAmendmentofExistingReport.Click();
					repo.MessageFromWebpage.ButtonOK.Click();
					break;
				case "Add Schedule A":
					repo.DomScotia.RequestService.AppraisalUpdateRequestType.AddRentalScheduleAtoExistingReport.Click();
					repo.MessageFromWebpage.ButtonOK.Click();
					break;
       			default:
        			    MessageBox.Show("Please specify update service type.");
        			break;	    
			}
			
			// Check Original Nas Number is required
			repo.DomScotia.RequestService.AppraisalUpdateRequestType.SearchOriNbrBtn.Click();
			if (repo.MessageFromWebpage.ButtonOKInfo.Exists())
			{	
				repo.MessageFromWebpage.ButtonOK.Click();
				Report.Log(ReportLevel.Info, "Information", "Mandatory Original request number message verified.");
			}else
			{	Report.Log(ReportLevel.Warn, "Warning", "Mandatory Original request number message not presented, please check.");}
			
			repo.DomScotia.RequestService.AppraisalUpdateRequestType.OrgReqNbr.PressKeys(oriNbr);
			repo.DomScotia.RequestService.AppraisalUpdateRequestType.SearchOriNbrBtn.Click();
			Delay.Milliseconds(100);
		}
		
		
		public void inputRequestInfo(string service, string contact, string tel1, string tel2, string tel3, string spTxt, string email)
		{
			if (service == "Drive-By" || service == "Desktop")
			{
				//Fill inContact Name and telphone number
				repo.DomScotia.RequestService.LoanAndContactInformation.AppContactName.PressKeys(contact);
				repo.DomScotia.RequestService.LoanAndContactInformation.Tel1.PressKeys(tel1);
				repo.DomScotia.RequestService.LoanAndContactInformation.Tel2.PressKeys(tel2);
				repo.DomScotia.RequestService.LoanAndContactInformation.Tel3.PressKeys(tel3);
				repo.DomScotia.RequestService.LoanAndContactInformation.PreferNumber1.Click();
				Delay.Milliseconds(100);
			}
			
			repo.DomScotia.RequestService.SpecialInstructions.SpText.PressKeys(spTxt);
			repo.DomScotia.RequestService.EmailNotifications.AdditionalEmail.PressKeys(email);
			Delay.Milliseconds(100);
			
			if(repo.DomScotia.RequestService.Payment.InvoiceInfo.Exists())
			{repo.DomScotia.RequestService.Payment.Invoice.Click();}
			
			repo.DomScotia.RequestService.SubmitRequest.PerformClick();
			
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
			string tel1 = random.Next(111, 999).ToString();
			string tel2 = random.Next(111, 999).ToString();
			string tel3 = random.Next(1111,9999).ToString();
			
			Login UsrLogin = new Login();
			UsrLogin.lauchScotia();
			UsrLogin.UserLogin(varUid, varPwd);
						
			//Request New Update Service
			repo.DomScotia.MainMenu.RequestService.Click();
			repo.DomScotia.RequestService.ResidentialUpdate.Click();
			Delay.Milliseconds(100);

			//Get Expected Service Type
			if (varUpdateType == "Appraisal Update")
				{
					string SQL1 = "SELECT a.app_request_nbr, service_residential FROM app_request a where username = '" + varUid + "' and status = 'Completed' and service_residential = 'Full-Service' order by app_request_nbr desc limit 1";
					DataTable dt1 = QueryDB.RunQuery(SQL1);
				
					oriNbr = dt1.Rows[0][0].ToString();
					service = dt1.Rows[0][1].ToString();
				}else if (varUpdateType == "Add Schedule A")
				{
					string SQL2 = "SELECT a.app_request_nbr, service_residential FROM app_request a where username = '" + varUid + "' and status = 'Completed' and service_residential not like '%Schedule A%' order by app_request_nbr desc limit 1";
					DataTable dt2 = QueryDB.RunQuery(SQL2);
				
					oriNbr = dt2.Rows[0][0].ToString();
					service = dt2.Rows[0][1].ToString();
				}
			
			//Order Appraisal Update service			
			selectUpdate(varUpdateType, oriNbr);
			Delay.Milliseconds(100);
			
			inputRequestInfo(service, contactName, tel1, tel2, tel3, spText, addiEmail);
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
