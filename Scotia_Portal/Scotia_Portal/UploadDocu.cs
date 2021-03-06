/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-05-12
 * Time: 5:15 PM
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

namespace Scotia_Portal
{
	/// <summary>
	/// Description of UploadDocu.
	/// </summary>
	[TestModule("122C7257-0311-47E8-846E-624D281FFEF6", ModuleType.UserCode, 1)]
	public class UploadDocu : ITestModule
	{
		public static Scotia_PortalRepository repo = Scotia_PortalRepository.Instance;
		
			static UploadDocu instance = new UploadDocu();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public UploadDocu()
		{
			// Do not delete - a parameterless constructor is required!
		}
		
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("EE042C86-1145-445F-AEF7-3581ED307180")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		#endregion
		
		private string uid = "appraiseruattest.nas.com";
    	private string pwd = "Tester#1";
    	
    	private string FilePath = @"c:\Upload_Files\SupportDocument.pdf\";
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		/// 
		public void UploadFile(string file)
		{
			repo.DomScotia.SearchResults.UploadSupportDocu.Click();
			Delay.Milliseconds(100);
			
			repo.DomScotia.SearchResults.FilePath.DoubleClick();
			Delay.Milliseconds(200);
			
			repo.ChooseFileToUpload.FileRouteText.PressKeys(file);
			Delay.Milliseconds(200);
			repo.ChooseFileToUpload.ButtonOpen.Click();
			Delay.Milliseconds(400);
			
			repo.DomScotia.SearchResults.UploadSubmit.Click();
			Delay.Milliseconds(600);
			
		}
		
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			/*
			Login UsrLogin = new Login();
			UsrLogin.lauchScotia();
			UsrLogin.UserLogin(uid, pwd);
			Delay.Milliseconds(100);
			*/
			
			SearchRequest search = new SearchRequest();
			search.search();
			search.findByNasNbr(varNasNbr);
			Delay.Milliseconds(300);
			
			UploadDocu upload = new UploadDocu();
			upload.UploadFile(FilePath);
			Delay.Milliseconds(300);
			
			//Verify Upload Files successfully
			Report.Log(ReportLevel.Success, "Validation", "Supporting Document has been successfully uploaded for: " + varNasNbr);   
			Validate.Exists(repo.DomScotia.SearchFilter.SupportingDocumentsHaveBeenUploaded);
			Delay.Milliseconds(200);
			
		}
	}
}
