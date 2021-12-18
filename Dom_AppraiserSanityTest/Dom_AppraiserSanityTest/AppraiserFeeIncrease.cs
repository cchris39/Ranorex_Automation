/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 26/01/2016
 * Time: 9:59 AM
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

namespace Dom_AppraiserSanityTest
{
	/// <summary>
	/// Description of AppraiserFeeIncrease.
	/// Scenario: Appraiser submit two fee increase requests: "Same day Service"
	/// </summary>
	[TestModule("855E187D-C9E6-4142-B2B7-F59B448F457C", ModuleType.UserCode, 1)]
	public class AppraiserFeeIncrease : ITestModule
	{
		/// <summary>
		/// Using the Dom_AppraiserSanityTestRepository repository.
		/// </summary>
		public static Dom_AppraiserSanityTestRepository repo = Dom_AppraiserSanityTestRepository.Instance;
		
		static AppraiserFeeIncrease instance = new AppraiserFeeIncrease();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public AppraiserFeeIncrease()
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
		
		string _varAppraiser = "";
		[TestVariable("53A10DCD-834B-4679-988C-ABD6435C41FA")]
		public string varAppraiser
		{
			get { return _varAppraiser; }
			set { _varAppraiser = value; }
		}
		
		string _varPwd = "";
		[TestVariable("BDE26ECC-69C8-43F7-9D0A-B8C0909E6391")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varUrl = "";
		[TestVariable("ACB9F9D1-9AA1-4C24-BADB-0C25001DE3C7")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("EA4FD2AB-3FCA-4279-8441-E4FA5FF53B1E")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
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
			
			//Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(300);
			
			//Appraiser log in
			//repo.DomNasHome.Userid.PressKeys(varAppraiser);   
			//repo.DomNasHome.Password.PressKeys(varPwd);
			//repo.DomNasHome.Submit.Click();
			//Delay.Milliseconds(100);
			
			//Search By Nas Number
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(varNasNbr);    
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			//Check request current status
			string feeAmount = "50.00";
			string curStatus = repo.DomNasHome.MenuDisplay.RequestStatus.InnerText.Trim();
			string feeReason = "Same Day Service";
			
		
			
			if (curStatus != "Completed")
			{
				repo.DomNasHome.MenuDisplay.FeeIncreaseBtn.Click();
				repo.DomNasHome.MenuDisplay.AddiFee1Desc.Element.SetAttributeValue("TagValue", feeReason);
				Delay.Milliseconds(100);
				repo.DomNasHome.MenuDisplay.AddiFee1.PressKeys(feeAmount.Trim());
				Delay.Milliseconds(100);
				
				repo.DomNasHome.MenuDisplay.FeeNotes.PressKeys("Fee increase test");
				repo.DomNasHome.MenuDisplay.FeeSubmitButton.Click();
				Delay.Milliseconds(100);
				
				string newFee = repo.DomNasHome.MenuDisplay.TotalNewFee.InnerText.Trim();
			}
			//Report Fee Submit Status;
			Report.Log(ReportLevel.Success,"Validation", "Request Fee increase submit successfully for:  " + varNasNbr);
			Validate.Exists(repo.DomNasHome.MenuDisplay.FeeIncreaseRecordCreatedSuccessfully);
			
			Delay.Milliseconds(100);
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
		}
	}
}
