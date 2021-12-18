/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 14/03/2016
 * Time: 4:33 PM
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
	/// Description of ChangeStatus.
	/// </summary>
	[TestModule("D6BC9AD2-7AF6-445A-ADF4-BF6E339D789C", ModuleType.UserCode, 1)]
	public class ChangeStatus : ITestModule
	{
		
		/// <summary>
		/// Using the Dom_SanityTestRepository repository.
		/// </summary>
		public static BrokerFlowRepository repo = BrokerFlowRepository.Instance;
		
		static ChangeStatus instance = new ChangeStatus();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ChangeStatus()
		{
			// Do not delete - a parameterless constructor is required!
		}
		
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("43FCD8F1-7053-4B89-BD9F-FED748E5EBB9")]
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
			
			//Search By Nas Number
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(varNasNbr);     // varNasNbr
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			//Get current status from search result
			var status = repo.DomNasHome.MenuDisplay.StrongTagStatus.InnerText.Trim();
			const string changeStatus = "Cancelled";
			
			//Change request status 
				if (status != "Completed") 
				{
				repo.DomNasHome.MenuDisplay.ChgStatusBtn.Click();
				repo.DomNasHome.MenuDisplay.ButtonTagYes.Click();
				Delay.Milliseconds(200);
				
				var chgStatus = repo.DomNasHome.MenuDisplay.StrongTagStatus.InnerText.Trim();
				
				//report change status
				Report.Log(ReportLevel.Success, "Validation", "Request has been successfully cancelled.");
				Report.Log(ReportLevel.Info, "Validation", varNasNbr + "Current status is: " + chgStatus);     //varNasNbr
				Validate.AreEqual(changeStatus, chgStatus);
				Delay.Milliseconds(100);
				}
				else if (status == "Completed")
				{
				Report.Log(ReportLevel.Info, "Warning", "Request status is completed, it can not be cancelled.");
				}
				else 
				{
				Report.Log(ReportLevel.Info, "Validation", varNasNbr + " " + "Current status is: " + status);     //varNasNbr
				Report.Log(ReportLevel.Failure, "Validation", "Request has not been cancelled.");
				Validate.NotExists(repo.DomNasHome.MenuDisplay.StatusChangedFromCreditPendingToCa);
				Delay.Milliseconds(100);
				}
		    
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
		}
	}
}
