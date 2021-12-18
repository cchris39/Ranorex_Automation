/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 25/01/2016
 * Time: 4:39 PM
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

namespace Dom_AppraiserSanityTest
{
	/// <summary>
	/// Description of AppraiserAppointment.
	/// Scenario: Appraiser set up appointment time at 3 days after today at 9:00 AM time
	/// </summary>
	[TestModule("80EB81F1-6504-4D01-A76D-0468B5A49695", ModuleType.UserCode, 1)]
	public class AppraiserAppointment : ITestModule
	{
		/// <summary>
		/// Using the Dom_AppraiserSanityTestRepository repository.
		/// </summary>
		public static Dom_AppraiserSanityTestRepository repo = Dom_AppraiserSanityTestRepository.Instance;
		
		static AppraiserAppointment instance = new AppraiserAppointment();
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public AppraiserAppointment()
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
		
		string _varAppraiser = "";
		[TestVariable("4C3BD9BC-2498-4A28-92D8-AFC6E491EEF1")]
		public string varAppraiser
		{
			get { return _varAppraiser; }
			set { _varAppraiser = value; }
		}
		
		string _varPwd = "";
		[TestVariable("EF441E20-BAA5-461C-9342-DAA6DD21FE64")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varUrl = "";
		[TestVariable("8D9B1936-69A8-4684-90B8-CA69F4CA492F")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("517EB339-AF9D-411D-ABC3-55BF497D5FE8")]
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
						
			//Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(300);
			
						
			//Appraiser log in
			//repo.DomNasHome.Userid.PressKeys(varAppraiser);
			//repo.DomNasHome.Password.PressKeys(varPwd);
			//repo.DomNasHome.Submit.Click();
			//Delay.Milliseconds(100);
			
			//Search By Nas Number
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(varNasNbr);     // varNasNbr
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			repo.DomNasHome.MenuDisplay.SetAppointmentTimeBtn.Click();
			Delay.Milliseconds(100);
			
			//Set appointment date 3 days after the current day
			var curDate = System.DateTime.Today;
			var appointDate = curDate.AddDays(5).ToString("yyyy-MM-dd");
						
			repo.DomNasHome.MenuDisplay.AppoinmentDate.TagValue = appointDate;
			Delay.Milliseconds(100);
			
			
			//Set appointment time to 9:00 AM 
			repo.DomNasHome.MenuDisplay.HourSelect.TagValue = "09";
			repo.DomNasHome.MenuDisplay.MinuteSelect.TagValue = "00";
			repo.DomNasHome.MenuDisplay.AmPmButton.Value = "AM";
			repo.DomNasHome.MenuDisplay.AppointmentSubmitBtn.Click();
			Delay.Milliseconds(100);
			
			//Report Appointment Made Status			
			Report.Log(ReportLevel.Success, "Validation", "Appointment made successfully for " + varNasNbr + " at: " + appointDate + " 09:00:AM");
			Validate.Exists(repo.DomNasHome.MenuDisplay.AppointmentMadeSuccessfully);
			
			//Close Browser
			//Host.Local.KillBrowser("IE");
			//Delay.Milliseconds(200);
			
			
			
		}
	}
}
