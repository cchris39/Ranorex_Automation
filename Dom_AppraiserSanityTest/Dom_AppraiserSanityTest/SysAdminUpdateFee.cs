/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 26/01/2016
 * Time: 10:57 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace Dom_AppraiserSanityTest
{
	/// <summary>
	/// Description of SysAdminUpdateFee.
	/// Scenario: System admin (nasadmin) update fee status to 'Approved' in order for the appraiser to complete/upload the report
	/// </summary>
	[TestModule("37167EE5-DCB3-443B-928E-864D14A9B18D", ModuleType.UserCode, 1)]
	public class SysAdminUpdateFee : ITestModule
	{
		/// <summary>
		/// Using the Dom_AppraiserSanityTestRepository repository.
		/// </summary>
		public static Dom_AppraiserSanityTestRepository repo = Dom_AppraiserSanityTestRepository.Instance;
		
		static SysAdminUpdateFee instance = new SysAdminUpdateFee();
		/// <summary>
		
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public SysAdminUpdateFee()
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
		[TestVariable("FD38293D-60E0-4201-94CA-7D0FF41F633B")]
		public string varAdmin
		{
			get { return _varAdmin; }
			set { _varAdmin = value; }
		}
		
		string _varPwd = "";
		[TestVariable("79705193-A1DB-4393-AD0A-2164B04D8A0A")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varUrl = "";
		[TestVariable("8C79B7EA-CE0F-4650-8799-D453DD7C53EA")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("17444ABC-DD15-4857-AA37-73FF3933C4C9")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varAppraiser = "";
		[TestVariable("D26AEEFE-2399-4F89-8170-B3772CBBDE7E")]
		public string varAppraiser
		{
			get { return _varAppraiser; }
			set { _varAppraiser = value; }
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
			repo.DomNasHome.Userid.PressKeys(varAdmin);
			repo.DomNasHome.Password.PressKeys(varPwd);
			repo.DomNasHome.RememberUserName.Click();
			repo.DomNasHome.Submit.Click();
			Delay.Milliseconds(100);
			
			repo.SavePwdNo.Click();
			Delay.Milliseconds(200);
			
			//Go to Fee Administration
			repo.DomNasHome.SystemAdmin.Click();
			repo.DomNasHome.FeeAdministration.Click();
			Delay.Milliseconds(100);
			
			repo.DomNasHome.FeeAdmin.FeeAdminNasReqNum.PressKeys(varNasNbr);
			//repo.DomNasHome.FeeAdmin.FeeAdminAppraiser.PressKeys(varAppraiser);
			repo.DomNasHome.FeeAdmin.FeeAdminSearchSubmit.Click();
			Delay.Milliseconds(100);
			
			string curFeeStatus = repo.DomNasHome.FeeAdmin.FeeAdminFeeStatus.InnerText.Trim();
			
			if (curFeeStatus == "Pending")
			{
				repo.DomNasHome.FeeAdmin.RFeeIncreaseID.Click();
				repo.DomNasHome.FeeAdmin.FeeAdminViewFeeDetails.Click();
				Delay.Milliseconds(200);
				//Update Fee Status to 'Approved'
				repo.DomNasHome.FeeAdmin.FeeAdminFeeStatusUpdateOpt.Element.SetAttributeValue("TagValue", "Approved");
				Delay.Milliseconds(100);
				repo.DomNasHome.FeeAdmin.FeeAdminUpdate.Click();
				Delay.Milliseconds(200);
				
				//Rerport Fee status update result
				Report.Log(ReportLevel.Success, "Validation", "Fee status successfully approved for: " + varNasNbr);
				Validate.Exists(repo.DomNasHome.FeeAdmin.FeeIncreaseHasBeenUpdatedSuccessfully);
			} else
			{
				Report.Log(ReportLevel.Failure, "Validation", "Current Fee status for: " + varNasNbr + " is " + curFeeStatus + "; No fee update been done");
			}
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(300);
			
		}
	}
}
