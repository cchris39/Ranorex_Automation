/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 03/02/2016
 * Time: 9:51 AM
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
using System.Collections;
using System.Web;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace NasAdmin
{
	/// <summary>
	/// Description of ReassignAppraiser.
	/// </summary>
	[TestModule("856CD72A-FBEF-452F-BE00-E6B5D95EF52B", ModuleType.UserCode, 1)]
	public class ReassignAppraiser : ITestModule
	{
		/// <summary>
		/// Using the NasAdminRepository repository.
		/// </summary>
		public static NasAdminRepository repo = NasAdminRepository.Instance;
		
		static ReassignAppraiser instance = new ReassignAppraiser();
		
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ReassignAppraiser()
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
		
		string _varAdmin = "";
		[TestVariable("A7FA5DD0-D0A0-4455-BF2D-C1F711B7CE4C")]
		public string varAdmin
		{
			get { return _varAdmin; }
			set { _varAdmin = value; }
		}
		
		string _varPwd = "";
		[TestVariable("AE3C4B97-2E76-4961-808E-DFB338690249")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("2F923259-33CC-47BA-9617-2BA62791D235")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varUrl = "";
		[TestVariable("12A58275-8EE6-4009-904A-9A4149F0CF99")]
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
		
		private string usrContact;
		private string firstName;
		private string lastName;
		private const string msgRes = "Please select the reassigning reason";
		private const string resReason = "Better TAT";
		
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			
			Delay.Milliseconds(200);
			//Search Nas Number
			repo.Dom_NasHome.NasAdminMenu.SearchFilter.Click();
			repo.Dom_NasHome.MenuDisplay.NasReqNumSearch.PressKeys(varNasNbr);
			repo.Dom_NasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			string searchNasNbr = repo.Dom_NasHome.MenuDisplay.DivTagNasNbr.InnerText.Trim();
			
			if (searchNasNbr != varNasNbr)
			{
				Report.Log(ReportLevel.Warn,"Warning", "Nas Number is not match, please check.");
			}
			repo.Dom_NasHome.NasAdminFunction.ReassignAppraiser.Click();
			Delay.Milliseconds(300);
			
			/* Reassign to first available appraiser */
			//Get Re-assign Appraiser table
			Ranorex.TableTag aprTbl = repo.Dom_NasHome.MenuDisplay.ReassignAppraiserTbl;
			
			//Find enable radio button (The first the available appraiser)
			InputTag selRadio = aprTbl.FindSingle(".//tr[]/td[1]/input[@disabled='False' and @type='radio']");
			
			if (selRadio == null)
			{
				Report.Info("No appraiser can be Reassign to.");
			}
			selRadio.Click();
			Delay.Milliseconds(100);
			repo.Dom_NasHome.MenuDisplay.ReassignAppraiserSubmitBtn.Click();
			
			//CHECK POP OP MESSAGE FOR NOT SELECTING RE-ASSIGN REASON
			string resMsg = repo.MessageFromWebpage.Text65535_Reassign.Caption.Trim();
			if (resMsg == msgRes)
			{
			Report.Log(ReportLevel.Success, "Success", "Select reassign reason message verified.");
			}else
			{Report.Log(ReportLevel.Warn, "Warning", "Verified select reassign reason message failed, please check.");}
			
			//Select Reassign Reason 
			repo.MessageFromWebpage.ButtonOK.Click();
			repo.Dom_NasHome.MenuDisplay.ReassignReason.TagValue = resReason;
			Delay.Milliseconds(100);
			
			//Submit Reassign request
			repo.Dom_NasHome.MenuDisplay.ReassignAppraiserSubmitBtn.Click();
			Delay.Milliseconds(400);
					
			//Get UID of Assign Appraiser
			string asgnAppraiser = repo.Dom_NasHome.MenuDisplay.ATagAssignAppraiser.InnerText.Trim();
			
			//varAppraiser = asgnAppraiser;
			const string confirm = "Reassigned: " + resReason;
			repo.Dom_NasHome.NasAdminFunction.AppraisalHistory.Click();
			Delay.Milliseconds(200);
			
			string rasgnConfirm = repo.Dom_NasHome.MenuDisplay.ReassignedConfirmed.InnerText.Trim();
			
			//Report Reassign status
			Report.Log(ReportLevel.Success,"Info", "Request has been successfully reassign to: " + asgnAppraiser);
			Validate.AreEqual(rasgnConfirm,confirm);
			
			
			Delay.Milliseconds(300);
			/* Reassign to specific appraiser*/
			//Close Browser
			
			//End of Main Run Section
		}
	}
}
