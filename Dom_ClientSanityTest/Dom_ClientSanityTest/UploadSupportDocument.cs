/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 06/01/2016
 * Time: 9:28 AM
 * Version 1.0
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

namespace Dom_ClientSanityTest
{
	/// <summary>
	/// Description of UploadSupportDocument.
	/// Scenario: Client upload support documents after submitting order
	/// </summary>
	[TestModule("C163D7DC-2434-46E9-AAC4-8EAFE744C1F2", ModuleType.UserCode, 1)]
	public class UploadSupportDocument : ITestModule
	{
		/// <summary>
		/// Using the Dom_SanityTestRepository repository.
		/// </summary>
		public static Dom_SanityTestRepository repo = Dom_SanityTestRepository.Instance;
			
		static UploadSupportDocument instance = new UploadSupportDocument();
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public UploadSupportDocument()
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
		
		string _varUid = "";
		[TestVariable("4A2FB55A-C9A5-4536-8568-A014FBCDDD5A")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		
		string _varPwd = "";
		[TestVariable("364EDE3B-C882-4402-BD23-5B403FA32D8B")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("C5E4BE9B-F346-45F5-ADB8-D6CCDCA1B432")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varUrl = "";
		[TestVariable("2EC2972F-A5EF-4393-AA5D-41539CA41989")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
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
			
			/*/
			Host.Local.ClearBrowserCookies("IE");
			Delay.Milliseconds(100);
			
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(200);
			
			//Login
			repo.DomNasHome.Userid.PressKeys(varUid);
			repo.DomNasHome.Password.PressKeys(varPwd);
			repo.DomNasHome.Submit.Click();
			Delay.Milliseconds(100);
			/*/
			
			Delay.Milliseconds(200);	
			//Search By Nas Number
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(varNasNbr);     // varNasNbr
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(200);
			
			//Upload Support documents
			repo.DomNasHome.MenuDisplay.UploadSupportDocuBtn.Click();
			
			
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
