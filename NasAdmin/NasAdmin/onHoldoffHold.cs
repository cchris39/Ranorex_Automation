/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 05/02/2016
 * Time: 11:40 AM
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

namespace NasAdmin
{
	/// <summary>
	/// Description of OnHoldoffHold.
	/// </summary>
	[TestModule("DBD134E3-43DD-4250-B197-49027175041E", ModuleType.UserCode, 1)]
	public class OnHoldoffHold : ITestModule
	{
		
		/// <summary>
		/// Using the NasAdminRepository repository.
		/// </summary>
		public static NasAdminRepository repo = NasAdminRepository.Instance;
		
		static Resubmit instance = new Resubmit();
		
		#region Varilables
		
		string _varUrl = "";
		[TestVariable("B59F7CB3-1C66-4BB6-9297-1833108585B7")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("37341607-D370-4A37-8063-98DA43864288")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		
		#endregion
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public OnHoldoffHold()
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
			
			const string holdApplicant = "No response from Applicant";    //OHA1
			const string holdAppraiser = "No reponse from Appraiser";     //OHR1
			const string holdClient = "No reponse from Client";          //OHC1
			
			
			//Search Nas Number
			repo.Dom_NasHome.NasAdminMenu.SearchFilter.Click();
			repo.Dom_NasHome.MenuDisplay.NasReqNumSearch.PressKeys(varNasNbr);      //varNasNbr
			repo.Dom_NasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(200);
			
			//On Hold  -- Delay by Applicant
			repo.Dom_NasHome.NasAdminFunction.OnHold.Click();
			Delay.Milliseconds(100);
			repo.Dom_NasHome.MenuDisplay.Hold.ReasonSelect.TagValue = "OHA1";     // "No response form Applicant"
			Delay.Milliseconds(300);
			
			string delayBy1 = repo.Dom_NasHome.MenuDisplay.Hold.DelayBy.TagValue.ToString().Trim();
			string day1 = repo.Dom_NasHome.MenuDisplay.Hold.OnHoldDays.TagValue.ToString().Trim();
			string hour1 = repo.Dom_NasHome.MenuDisplay.Hold.OnHoldHours.TagValue.ToString().Trim();
			string minute1 = repo.Dom_NasHome.MenuDisplay.Hold.OnHoldMinutes.TagValue.ToString().Trim();
			string holdComment1 = repo.Dom_NasHome.MenuDisplay.Hold.HoldCommentText.InnerText.ToString().Trim();
			
			string holdTime1 = day1 + hour1 + minute1;
			
			repo.Dom_NasHome.MenuDisplay.Hold.HoldSubmit.Click();
			Delay.Milliseconds(100);
			
			repo.MessageFromWebpage.ButtonOK.Click();
			
			Report.Log(ReportLevel.Success, "Validation", "Request " + varNasNbr + " has been pout on hold for: " + holdTime1);
			Validate.Exists(repo.Dom_NasHome.MenuDisplay.Hold.AppraisalRequestNumberHasBeenOffHold);
			
			Report.Log(ReportLevel.Info, "Information", "Request on hold information is match.");
			Validate.AreEqual(holdApplicant, holdComment1);
			Validate.AreEqual(delayBy1, "Applicant");
			Delay.Milliseconds(100);
						
			//Check Follow Up Filter for the request being on hold
			repo.Dom_NasHome.NasAdminMenu.SystemAdmin.Click();
			Delay.Milliseconds(100);
			repo.Dom_NasHome.NasAdminMenu.FollowUpFilter.Click();
			
			repo.Dom_NasHome.MenuDisplay.FollowUp.PoNbr.PressKeys(varNasNbr);
			repo.Dom_NasHome.MenuDisplay.Submit.Click();
			Delay.Milliseconds(100);
			
			//Request found on Follow op filter
			Report.Log(ReportLevel.Success, "Validation", "On Hold Request " + varNasNbr + " is verified on Follow Up Filter.");
			Validate.Exists(repo.Dom_NasHome.MenuDisplay.PTag1RecordSFound);
			
			string nasNbr_1 = repo.Dom_NasHome.MenuDisplay.Hold.FontTagNasNbr.InnerText.Trim();
			Validate.AreEqual(nasNbr_1, varNasNbr);
			
			//Put request Off Hold
			repo.Dom_NasHome.MenuDisplay.Hold.OffHold_Radio.Click();
			Delay.Milliseconds(100);
			repo.Dom_NasHome.MenuDisplay.Hold.OffHoldBtn.Click();
			Delay.Milliseconds(100);
			repo.Dom_NasHome.NasAdminFunction.Submit.Click();
			
			//Report Off Hold Status
			Report.Log(ReportLevel.Success, "Validation", "Request " + varNasNbr + " has been successfully taken off hold.");
			Validate.Exists(repo.Dom_NasHome.NasAdminFunction.OffHoldConfirmed);
			Delay.Milliseconds(100);
			
			//Close "PutRequestOnOffTab"
			repo.OffHoldConfirmedTab.CloseTabCtrlPlusW.Click();
			Delay.Milliseconds(300);
			
			
			//Close Browser
			//Host.Local.KillBrowser("IE");
			//Delay.Milliseconds(300);
				
			
			
		}
	}
}
