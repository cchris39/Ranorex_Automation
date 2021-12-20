/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-04-29
 * Time: 2:50 PM
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
	/// Description of ChangeStatus.
	/// </summary>
	[TestModule("2D50C239-75CA-4F58-BDF5-F8C6AD76DEB4", ModuleType.UserCode, 1)]
	public class ChangeStatus : ITestModule
	{
		public static RBC_TestRepository repo = RBC_TestRepository.Instance;
		
		static ChangeStatus instance = new ChangeStatus();
    		
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("619FF5CE-377D-406F-B335-BFC9076BCA1C")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		#endregion
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ChangeStatus()
		{
			// Do not delete - a parameterless constructor is required!
		}

		public void CancellRequest()
		{
			repo.RBC_Portal.Main.Status_Cancell.Click();
			Delay.Milliseconds(100);
			
			if(repo.MessageFromWebpage.Text65535Info.Exists())
			{repo.MessageFromWebpage.ButtonOK.Click();}
			Delay.Milliseconds(300);
		}
		
		public void ChangeToReviewed()
		{
			repo.RBC_Portal.Main.Status_Reviewed.Click();
			Delay.Milliseconds(100);
			
			Report.Log(ReportLevel.Success, "Request " + varNasNbr + " was changed to reviewed successfully.");
			Validate.Exists(repo.RBC_Portal.Main.DivTagStatusChanged);
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
			
			SearchRequest search = new SearchRequest();
			search.findByNasNbr(varNasNbr);
			Delay.Milliseconds(300);
			
			repo.RBC_Portal.SearchRequest.SelectRadioRow.Click();
			Delay.Milliseconds(200);
			
			ChangeStatus cancell = new ChangeStatus();
			cancell.CancellRequest();
			Delay.Milliseconds(200);
			
			
			//Verify Status change status
			Report.Log(ReportLevel.Success, "Request " + varNasNbr + " was cancelled successfully.");
			Validate.Exists(repo.RBC_Portal.Main.DivTagStatusChanged);
			
			
			Delay.Milliseconds(200);
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
		}
	}
}
