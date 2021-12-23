/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-05-12
 * Time: 5:30 PM
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
	/// Description of ChangeStatus.
	/// </summary>
	[TestModule("0276F82F-CFE7-4D2F-914C-EFBB1D0354AF", ModuleType.UserCode, 1)]
	public class ChangeStatus : ITestModule
	{
		
		public static Scotia_PortalRepository repo = Scotia_PortalRepository.Instance;
		
			static ChangeStatus instance = new ChangeStatus();
		
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ChangeStatus()
		{
			// Do not delete - a parameterless constructor is required!
		}
		
		#region Varilabes
		
		string _varNasNbr = "";
		[TestVariable("A4DF8CC0-F154-4F6D-97FC-B129FD9237E0")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varStatus = "";
		[TestVariable("15C64BE3-42DC-4324-824E-16563B22F902")]
		public string varStatus
		{
			get { return _varStatus; }
			set { _varStatus = value; }
		}
		
		#endregion
		
		public void cancellRequest()
		{
			repo.DomScotia.SearchResults.ChagStatusCancel.Click();
			Delay.Milliseconds(100);
			repo.DomScotia.SearchResults.ConfirmYes.Click();
			Delay.Milliseconds(300);
			
			string curStatus = repo.DomScotia.SearchResults.curStatus.InnerText.Trim();
			Delay.Milliseconds(100);
						
			Report.Log(ReportLevel.Success, "Validation", "Request status has been successfully changed form " + varStatus + " to: " +  curStatus);
			Validate.Exists(repo.DomScotia.SearchResults.StatusChangedTag);
			
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
			
			Delay.Milliseconds(100);
			
			repo.DomScotia.SearchFilter.NasReqNum.PressKeys(varNasNbr);
			repo.DomScotia.SearchFilter.ViewGroupReq.Click();
			repo.DomScotia.SearchFilter.SearchSubmit.Click();
			Delay.Milliseconds(200);
			
			cancellRequest();
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
			
			
		}
	}
}
