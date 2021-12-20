/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 05/02/2016
 * Time: 9:58 AM
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

namespace NasAdmin
{
	/// <summary>
	/// Description of Resubmit.
	/// </summary>
	[TestModule("05CFDFCE-33FB-482F-BD38-0E8190FEC16D", ModuleType.UserCode, 1)]
	public class Resubmit : ITestModule
	{
		
		/// <summary>
		/// Using the NasAdminRepository repository.
		/// </summary>
		public static NasAdminRepository repo = NasAdminRepository.Instance;
		
		static Resubmit instance = new Resubmit();
		
		#region Varilables
		
		string _varAdmin = "";
		[TestVariable("36678527-9F22-4DD6-8474-D90149D3E18F")]
		public string varAdmin
		{
			get { return _varAdmin; }
			set { _varAdmin = value; }
		}
		
		string _varPwd = "";
		[TestVariable("A948A3E4-22F2-43A4-AD32-DD4BBCA57630")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varUrl = "";
		[TestVariable("189D9BAC-A800-4E24-87D5-9C0B6097EB1F")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("0AF1BB56-3114-4B75-B787-B65C8C35ED06")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		#endregion
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public Resubmit()
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
			Delay.Milliseconds(0);
			
			//Admin User log in
			repo.Dom_NasHome.Userid.PressKeys(varAdmin);
			repo.Dom_NasHome.Password.PressKeys(varPwd);
			repo.Dom_NasHome.RememberUserName.Click();
			repo.Dom_NasHome.LoginSubmit.Click();
			Delay.Milliseconds(300);
			
			//repo.NotForThisSite.Click();
			
			//WebDocument DomWeb = NasRepo.Dom_NasHome.Self; 
			
			//Search Nas Number
			repo.Dom_NasHome.NasAdminMenu.SearchFilter.Click();
			repo.Dom_NasHome.MenuDisplay.NasReqNumSearch.PressKeys(varNasNbr);      //varNasNbr
			repo.Dom_NasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(200);
			
			//Define new request varilables and assign new value
			string NewStrNum = "369";
			string NewStrName = "Tree House";
			string NewStrType = "St";
			string NewPoCode = "V3K 1P0";
			string NewAppName = "Westwood Creek";
			string NewSerType = "Drive-By";
			
			string NewAddress = NewStrNum + " " + NewStrName + " " + NewStrType;
			
			//Resubmit Request
			repo.Dom_NasHome.NasAdminFunction.ResubmitRequest.Click();
			Delay.Milliseconds(100);
			
			//Change Street number
			repo.Dom_NasHome.MenuDisplay.ResubmitStreetNum.Value = NewStrNum;
			
			//Change Street Name
			repo.Dom_NasHome.MenuDisplay.ResubmitAppStreetName.Value = NewStrName;
			Delay.Milliseconds(100);
			
			//change STREET TYPE
			repo.Dom_NasHome.MenuDisplay.ResubmitStreetType.Value = NewStrType;
			
			//change Postal Code
			repo.Dom_NasHome.MenuDisplay.ResubmitPostalCode.Value = NewPoCode;
			Delay.Milliseconds(100);
			
			//change Applicant Name
			repo.Dom_NasHome.MenuDisplay.ResubmitAppName.Value = NewAppName;
			Delay.Milliseconds(100);
			
			//Change Service Type
			repo.Dom_NasHome.MenuDisplay.ResubmitServiceResidential.TagValue = NewSerType;
			
			//Change Special Direction Text
			repo.Dom_NasHome.MenuDisplay.ResubmitSpText.TagValue = "Request information changed, please be aware. Thanks";
			Delay.Milliseconds(100);
			
			repo.Dom_NasHome.MenuDisplay.Resubmit.Click();
			Delay.Milliseconds(200);
			 
			Report.Log(ReportLevel.Success, "Validation", "Request " + varNasNbr + "  has been successfully resubmited." );
			Validate.Exists(repo.Dom_NasHome.MenuDisplay.UpdateSuccessfully);
			
			//Query DB to verified the New Information 
			string SQL = "SELECT a.`app_address`, a.`app_postal_code`, a.`app_name`, a.`service_residential` FROM app_request a where app_request_nbr = " + varNasNbr + ";";
			
			DataTable dt = DatabaseQuery.RunQuery(SQL);
			
			string strAddress = dt.Rows[0][0].ToString().Trim();
			string poCode = dt.Rows[0][1].ToString().Trim();
			string appName = dt.Rows[0][2].ToString().Trim();
			string serType = dt.Rows[0][3].ToString().Trim();
			
			//Verified the new information against the input
			Report.Log(ReportLevel.Success, "Validation", "Request " + varNasNbr + "  property address updated successfully." );
			Validate.AreEqual(strAddress, NewAddress);
			Delay.Milliseconds(100);
			
			Report.Log(ReportLevel.Success, "Validation", "Request " + varNasNbr + "  postal code updated successfully." );
			Validate.AreEqual(poCode, NewPoCode);
			Delay.Milliseconds(100);
			
			Report.Log(ReportLevel.Success, "Validation", "Request " + varNasNbr + "  applicant name updated successfully." );
			Validate.AreEqual(appName, NewAppName);
			Delay.Milliseconds(100);
			
			Report.Log(ReportLevel.Success, "Validation", "Request " + varNasNbr + "  service type updated successfully." );
			Validate.AreEqual(serType, NewSerType);
			Delay.Milliseconds(300);
			
			
			
		}
	}
}
