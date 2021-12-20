/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 16/02/2016
 * Time: 11:51 AM
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

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace NasAdmin
{
	/// <summary>
	/// Description of SetAppointment.
	/// </summary>
	[TestModule("4B7461C3-37CA-4BD6-9370-E927F5D2B200", ModuleType.UserCode, 1)]
	public class SetAppointment : ITestModule
	{
		/// <summary>
		/// Using the NasAdminRepository repository.
		/// </summary>
		public static NasAdminRepository repo = NasAdminRepository.Instance;
		
		static SetAppointment instance = new SetAppointment();
		
		#region Varilabels
		
		string _varNasNbr = "";
		[TestVariable("F9346EAF-ABBB-4D36-92DD-9A150FDC42C7")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varAdmin = "";
		[TestVariable("C3FD1AC4-3584-4342-A971-4109DD75226A")]
		public string varAdmin
		{
			get { return _varAdmin; }
			set { _varAdmin = value; }
		}
		
		string _varPwd = "";
		[TestVariable("49137A23-31F3-43E3-9DB4-8D8ED937EF22")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varUrl = "";
		[TestVariable("63E34A70-2E89-47B7-BCE6-A26E0955D35F")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		#endregion
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public SetAppointment()
		{
			// Do not delete - a parameterless constructor is required!
		}
		
		public void KillBrowser(string IE)
		{
		}

		public void ClearBrowserCookies(string IE)
		{
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
			
			/*/
			Host.Local.ClearBrowserCookies("IE");
			Delay.Milliseconds(100);
			
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(0);
			
			//Admin User log in
			repo.Dom_NasHome.Userid.PressKeys(varAdmin);
			repo.Dom_NasHome.Password.PressKeys(varPwd);
			repo.Dom_NasHome.RememberUserName.Click();
			repo.Dom_NasHome.LoginSubmit.Click();
			Delay.Milliseconds(300);
			/*/
			
			repo.Dom_NasHome.NasAdminMenu.SearchFilter.Click();
			Delay.Milliseconds(100);	
			
			//Input Nas Number to search
			repo.Dom_NasHome.MenuDisplay.NasReqNumSearch.PressKeys(varNasNbr);
			repo.Dom_NasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			//Check request status
			string status = repo.Dom_NasHome.MenuDisplay.ReqStatus.InnerText.Trim();
			
			if (status == "Accepted")
			{
				repo.Dom_NasHome.NasAdminFunction.SetAppointment.Click();
				Delay.Milliseconds(100);
			}else if(status == "New")
			{
				//Accepted Request
				repo.Dom_NasHome.NasAdminFunction.Chastatus.TagValue = "Accepted";
				repo.Dom_NasHome.NasAdminFunction.StatusChangeBtn.Click();
				Delay.Milliseconds(200);
				
				repo.Dom_NasHome.NasAdminFunction.SetAppointment.Click();
			}else
			{
				Report.Warn ("Request " + varNasNbr + " is in " + status + "; No appointment time can be set. Please check.");
				Delay.Milliseconds(100);
			}
			
			//Set appointment time
			string apptDate = System.DateTime.Today.AddDays(5).ToString("yyyy-MM-dd");
			string apptHour = "11";
			string apptMin = "30";
			string apptAmPm = "AM";
			string apptTime = apptDate + " " + apptHour + ":" + apptMin + " " + apptAmPm;
			
			string nasNbr = repo.Dom_NasHome.MenuDisplay.Appointment.AppointmentAppraisalRequestNumber.InnerText.Trim();
			
			if (nasNbr == varNasNbr)
			{
				repo.Dom_NasHome.MenuDisplay.Appointment.AppoinmentDate.Value = apptDate;
				Delay.Milliseconds(200);
				repo.Dom_NasHome.MenuDisplay.Appointment.HourSelect.TagValue = apptHour;
				repo.Dom_NasHome.MenuDisplay.Appointment.MinuteSelect.TagValue = apptMin;
				repo.Dom_NasHome.MenuDisplay.Appointment.AmPmOption.Value = apptAmPm;
				Delay.Milliseconds(100);
				
				repo.Dom_NasHome.MenuDisplay.Appointment.AppointmentSubmitBtn.Click();
				Delay.Milliseconds(200);
				
				//Report Set appointment status
				Report.Log(ReportLevel.Success, "Validation", "Appointmetn time has been successfully set for " + varNasNbr + " at: " + apptTime);
				Validate.Exists(repo.Dom_NasHome.MenuDisplay.Appointment.AppointmentMadeSuccessfully);
			}else
			{
				Report.Log(ReportLevel.Info, "Validation", "Nas number " + varNasNbr +  " is not match; Please check your request number.");
				
			}
			
			//Close Browser
			//Host.Local.KillBrowser("IE");
			Delay.Milliseconds(100);
		}
	}
}
