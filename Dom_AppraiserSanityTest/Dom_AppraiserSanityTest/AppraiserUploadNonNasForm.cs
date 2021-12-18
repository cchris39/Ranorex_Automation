/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 26/01/2016
 * Time: 11:35 AM
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
using System.IO;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace Dom_AppraiserSanityTest
{
	/// <summary>
	/// Description of AppraiserUploadNonNasForm.
	/// Scenario: Appraiser upload a Non-nas Form with the active request
	/// </summary>
	[TestModule("8BE851D2-23AC-4B6D-ACC8-F6C7B272384A", ModuleType.UserCode, 1)]
	public class AppraiserUploadNonNasForm : ITestModule
	{
		/// <summary>
		/// Using the Dom_AppraiserSanityTestRepository repository.
		/// </summary>
		public static Dom_AppraiserSanityTestRepository repo = Dom_AppraiserSanityTestRepository.Instance;
		
		static AppraiserUploadNonNasForm instance = new AppraiserUploadNonNasForm();
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public AppraiserUploadNonNasForm()
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
		[TestVariable("DA4280D1-B625-401A-BF10-E65F96737C87")]
		public string varAppraiser
		{
			get { return _varAppraiser; }
			set { _varAppraiser = value; }
		}
		
		string _varPwd = "";
		[TestVariable("CA558DB8-B397-4A91-B45D-12B05E63163E")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varUrl = "";
		[TestVariable("300ADF94-CD7A-4A5F-B33C-F676209FA8F5")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("FBC13C39-1300-432F-86C7-A6B9B238ED86")]
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
			
			Host.Local.ClearBrowserCookies("IE");
			Delay.Milliseconds(200);
			
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(300);
			
			repo.DomNasHome.Userid.Element.SetAttributeValue ("TagVlaue", null);
			Delay.Milliseconds(100);
			//Appraiser log in
			repo.DomNasHome.Userid.PressKeys(varAppraiser);
			repo.DomNasHome.Password.PressKeys(varPwd);
			repo.DomNasHome.RememberUserName.Click();
			repo.DomNasHome.Submit.Click();
			Delay.Milliseconds(100);
			
			//Search By Nas Number
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(varNasNbr);     // varNasNbr
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			string curStatus = repo.DomNasHome.MenuDisplay.RequestStatus.InnerText.Trim();
			
			if (curStatus == "Completed")
			{
				Report.Log(ReportLevel.Info, "Warning", "Request " + varNasNbr + " has been completed. No report upload this time.");
			}
			
			//Change Status to completed
			repo.DomNasHome.MenuDisplay.Chastatus.Element.SetAttributeValue("TagValue", "Completed");
			repo.DomNasHome.MenuDisplay.ChgStatusBtn.Click();
			Delay.Milliseconds(100);
				
			repo.DomNasHome.MenuDisplay.UploadFileBrowse.DoubleClick();
			Delay.Milliseconds(100);
			
			string fileOld = @"c:\Upload_Files\toUpload.pdf";    
			string fileNew = fileOld.Replace("toUpload", varNasNbr);
			File.Copy(fileOld, fileNew);
			
			repo.ChooseFileToUpload.FileName.PressKeys(fileNew);
			Keyboard.Press("{Return}");
			Delay.Milliseconds(300);
			
			repo.DomNasHome.MenuDisplay.UploadFileBtn.Click();
			Delay.Milliseconds(800);
				
			//Verify Upload Files successfully
			Report.Log(ReportLevel.Success, "Validation", "Nas Report has been successfully uploaded for: " + varNasNbr);    //varNasNbr
			Validate.Exists(repo.DomNasHome.MenuDisplay.YourReportHasBeenSuccessfullyComple);
			Delay.Milliseconds(200);
				
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
				
			
		}
	}
}
