/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 14/03/2016
 * Time: 2:03 PM
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

namespace BrokerFlow
{
	/// <summary>
	/// Description of UploadSupportDocu.
	/// </summary>
	[TestModule("379079EE-2205-4106-913E-1D7769CDB493", ModuleType.UserCode, 1)]
	public class UploadSupportDocu : ITestModule
	{
		/// <summary>
		/// Using the Dom_SanityTestRepository repository.
		/// </summary>
		public static BrokerFlowRepository repo = BrokerFlowRepository.Instance;
		
		static UploadSupportDocu instance = new UploadSupportDocu();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public UploadSupportDocu()
		{
			// Do not delete - a parameterless constructor is required!
		}
		public void KillBrowser(string IE)
		{
		}

		public void ClearBrowserCookies(string IE)
		{
		}
		
		#region Variables
		string _varNasNbr = "";
		[TestVariable("71E63976-9450-490B-87D5-792BB51305A8")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varUrl = "";
		[TestVariable("7CA45DDB-B80E-4891-B7F0-1A7A642C1CFE")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varUid = "";
		[TestVariable("9205FC0A-45CF-41E1-8819-D234BF7C81C3")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("EA222993-F1F5-43FA-8DA5-94CE433EFC84")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
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
			
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(200);
			
			//Login
			repo.DomNasHome.Userid.PressKeys(varUid);
			repo.DomNasHome.Password.PressKeys(varPwd);
			repo.DomNasHome.Submit.Click();
			Delay.Milliseconds(200);
			
			//Search By Nas Number
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(varNasNbr);     // varNasNbr
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(200);
			
			//Upload Support documents
			repo.DomNasHome.MenuDisplay.UploadSupportDocuBtn.Click();
			//repo.DomNasHome.MenuDisplay.BrowseFileToUpload.Click;

			//Use Ranorex Xpath instead of repository object
			Ranorex.InputTag my_Upload = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//form[#'form1']/table[@innertext='   ']/tbody/tr[3]/?/?/input[@id='file']";

			string FilePath = @"c:\Upload_Files\SupportDocument.pdf\";
			// "c:\Upload_Files\TestDocument.doc\";    //"c:\Upload_Files\SupportDocument.pdf\"    //"c:\Upload_Files\TestHousePic.jpg\"
			
			
			my_Upload.DoubleClick();
			Delay.Milliseconds(100);
			repo.ChooseFileToUpload.FileName.PressKeys(FilePath);
			Keyboard.Press("{Return}");
			Delay.Milliseconds(600);
				
    		repo.DomNasHome.MenuDisplay.UploadSubmit.Click();
    		Delay.Milliseconds(600);
			
			//Verify Upload Files successfully
			Report.Log(ReportLevel.Success, "Validation", "Supporting Document has been successfully uploaded for: " + varNasNbr);    //varNasNbr
			Validate.Exists(repo.DomNasHome.MenuDisplay.SupportingDocumentsHaveBeenUploaded);
			Delay.Milliseconds(200);
				
			//Close Browser
			//Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
		}
	}
}
