/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-04-29
 * Time: 2:41 PM
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

namespace RBC_Test
{
	/// <summary>
	/// Description of UploadDocu.
	/// </summary>
	[TestModule("36D3310C-9A89-44F1-A6AD-87722F58A7F8", ModuleType.UserCode, 1)]
	public class UploadDocu : ITestModule
	{
		public static RBC_TestRepository repo = RBC_TestRepository.Instance;
		
		static UploadDocu instance = new UploadDocu();
		
		
		private string FilePath = @"c:\Upload_Files\SupportDocument.pdf\";
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public UploadDocu()
		{
			// Do not delete - a parameterless constructor is required!
		}
		
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("CD551AC3-B076-45CB-91AB-CE5882116135")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		#endregion
		
		public void UploadFile(string file)
		{
			repo.RBC_Portal.Main.UploadSupportDocument.Click();
			Delay.Milliseconds(100);
			
			repo.RBC_Portal.Main.FilePath.DoubleClick();
			Delay.Milliseconds(200);
			
			repo.ChooseFileToUpload.FilePath.PressKeys(file);
			Delay.Milliseconds(200);
			repo.ChooseFileToUpload.ButtonOpen.Click();
			Delay.Milliseconds(400);
			
			repo.RBC_Portal.Main.FileUpload.Click();
			Delay.Milliseconds(600);
			
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
			
			SearchRequest Search = new SearchRequest();
			Search.findByNasNbr(varNasNbr);
			Delay.Milliseconds(300);
			
			UploadDocu upload = new UploadDocu();
			upload.UploadFile(FilePath);
			Delay.Milliseconds(200);
			
			//Verify Upload Files successfully
			Report.Log(ReportLevel.Success, "Validation", "Supporting Document has been successfully uploaded for: " + varNasNbr);    //varNasNbr
			Validate.Exists(repo.RBC_Portal.Main.SupportingDocumentsHaveBeenUploaded);
			Delay.Milliseconds(100);
		
		}
	}
}
