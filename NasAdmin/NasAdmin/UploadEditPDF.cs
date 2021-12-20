/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 05/02/2016
 * Time: 6:08 PM
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

namespace NasAdmin
{
	/// <summary>
	/// Description of UploadEditPDF.
	/// </summary>
	[TestModule("FC343B6E-3706-4018-AB5E-35794A95423E", ModuleType.UserCode, 1)]
	public class UploadEditPDF : ITestModule
	{
		/// <summary>
		/// Using the NasAdminRepository repository.
		/// </summary>
		public static NasAdminRepository repo = NasAdminRepository.Instance;
		
		static UploadEditPDF instance = new UploadEditPDF();
		
		public void KillBrowser(string IE)
		{
		}

		public void ClearBrowserCookies(string IE)
		{
		}
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("ADB4E14B-4F5E-4230-8D14-6A26F39667B3")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varUrl = "";
		[TestVariable("3EA9C3C8-7339-4A1C-84C7-6BFDF2BAD254")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		#endregion
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public UploadEditPDF()
		{
			// Do not delete - a parameterless constructor is required!
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
						
					
			//WebDocument DomWeb = NasRepo.Dom_NasHome.Self; 

			//Search Nas Number
			repo.Dom_NasHome.NasAdminMenu.SearchFilter.Click();
			repo.Dom_NasHome.MenuDisplay.NasReqNumSearch.PressKeys(varNasNbr);      //varNasNbr
			repo.Dom_NasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(200);
			
			//Accepted Request
			repo.Dom_NasHome.NasAdminFunction.Chastatus.TagValue = "Accepted";
			repo.Dom_NasHome.NasAdminFunction.StatusChangeBtn.Click();
			Delay.Milliseconds(200);
			
			//Complete Request
			repo.Dom_NasHome.NasAdminFunction.Chastatus.TagValue = "Completed";
			repo.Dom_NasHome.NasAdminFunction.StatusChangeBtn.Click();
			Delay.Milliseconds(200);
			
			repo.Dom_NasHome.MenuDisplay.FileNameBrowse.DoubleClick();
			Delay.Milliseconds(100);
			
			
			string fileOld = @"c:\Upload_Files\toUpload.pdf";    
			string fileNew = fileOld.Replace("toUpload", varNasNbr);       //varNasNbr
			File.Copy(fileOld, fileNew);
			
			
			repo.ChooseFileToUpload.FilePath.PressKeys(fileNew);
			Keyboard.Press("{Return}");
			Delay.Milliseconds(300);
			
			repo.Dom_NasHome.MenuDisplay.UploadFileBtn.Click();
			Delay.Milliseconds(800);
			
			//Report Uplaod Status
			Report.Log(ReportLevel.Success, "Validation", "Report has been successfully uploaded for: " + varNasNbr);    //varNasNbr
			Validate.Exists(repo.Dom_NasHome.MenuDisplay.YourReportHasBeenSuccessfullyComple);
			Delay.Milliseconds(200);				
			
			//Edit PDF
			//Search Nas Number
			repo.Dom_NasHome.NasAdminMenu.SearchFilter.Click();
			repo.Dom_NasHome.MenuDisplay.NasReqNumSearch.PressKeys(varNasNbr);      //varNasNbr
			repo.Dom_NasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(200);
			
			repo.Dom_NasHome.NasAdminFunction.EditPDF.Click();
			Delay.Milliseconds(200);	
			
			repo.Dom_NasHome.MenuDisplay.FileNameBrowse.DoubleClick();
			Delay.Milliseconds(100);
						
			repo.ChooseFileToUpload.FilePath.PressKeys(fileNew);
			Keyboard.Press("{Return}");
			Delay.Milliseconds(300);
			
			repo.Dom_NasHome.MenuDisplay.UploadFileBtn.Click();
			Delay.Milliseconds(800);
			
			//Report Edit PDF Upload Status
			Report.Log(ReportLevel.Success, "Validation", "Report has been successfully re-uploaded for: " + varNasNbr);    //varNasNbr
			Validate.Exists(repo.Dom_NasHome.MenuDisplay.H2TagFileUploadedSuccessful);
			Delay.Milliseconds(200);	
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
			
			
		}
	}
}
