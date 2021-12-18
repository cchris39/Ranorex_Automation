/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 24/12/2015
 * Time: 9:56 AM
 * Version: 1.0
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

namespace Dom_ClientSanityTest
{
	/// <summary>
	/// Description of ViewAppHistory.
	/// Scenario: User view Appraisal History by searching Nas Number and verify the Appraisal History information
	/// </summary>
	[TestModule("B34D8D3F-7BC4-4056-906C-F74EB777E40E", ModuleType.UserCode, 1)]
	public class ViewAppHistory : ITestModule
	{
		#region Varilables
		string _varUid = "";
		[TestVariable("B77FC1A9-6263-4CC3-8058-617CB7483DBB")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("7A047383-1D9E-4066-A00C-495CA7FE755C")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("3AD647AE-1AB8-462C-9AED-96419C772922")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varCreationDate = "";
		[TestVariable("5D753D81-2C60-4864-B3D0-31871C6514B5")]
		public string varCreationDate
		{
			get { return _varCreationDate; }
			set { _varCreationDate = value; }
		}
		
		string _varUrl = "";
		[TestVariable("9A14FD75-C627-495E-AA8E-AE026F46EDE0")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		#endregion
		
		/// <summary>
		/// Using the Dom_SanityTestRepository repository.
		/// </summary>
		public static Dom_SanityTestRepository repo = Dom_SanityTestRepository.Instance;
		
		static ViewAppHistory instance = new ViewAppHistory();
				
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ViewAppHistory()
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
			
			
			/*/
			Host.Local.ClearBrowserCookies("IE");
			Delay.Milliseconds(100);
			
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(100);
			
			//Login
			repo.DomNasHome.Userid.PressKeys(varUid);
			repo.DomNasHome.Password.PressKeys(varPwd);
			repo.DomNasHome.Submit.Click();
			Delay.Milliseconds(100);
			/*/
			
			Delay.Milliseconds(100);
			//Search By Nas Number
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(varNasNbr);       //varNasNbr
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			repo.DomNasHome.MenuDisplay.AppraisalHistoryBtn.Click();
			
			string varCreateDate = repo.DomNasHome.MenuDisplay.TdTagTimeStamp.InnerText.Trim().Substring(0,10);
			string status = repo.DomNasHome.MenuDisplay.TdTagStatus.InnerText.Trim();
			string appPrice = repo.DomNasHome.MenuDisplay.TdTagAppPrice.InnerText.Trim();
			string appPriceTax = repo.DomNasHome.MenuDisplay.TdTagAppPriceTax.InnerText.Trim();
			string nasNbr = repo.DomNasHome.MenuDisplay.AppraisalRequestNumber.InnerText.Trim().Substring(26,7);
			string regionType =repo.DomNasHome.MenuDisplay.RegionType.InnerText.Trim().Substring(13,5);
			
			varCreationDate = varCreateDate;
			
			//Report Appraisal History Status and validate Status, App Price etc.,
			Report.Log(ReportLevel.Info, "Validation", "Appraisal History page present.");
			Validate.Exists(repo.DomNasHome.MenuDisplay.AppraisalRequestHistoryReport);
			Report.Log(ReportLevel.Success, "Validation", "Nas number: " + varNasNbr  + " match");
			Validate.AreEqual(nasNbr, varNasNbr);
			Report.Log(ReportLevel.Info, "Validation", "Region Type is: " + regionType);
			
			Report.Log(ReportLevel.Success, "Validation", "Request status is expected: " + status);
			Validate.AreEqual(status, "New");
				
			Report.Log(ReportLevel.Success, "Validation", "App Price is expected: " + appPrice);
			Validate.AreEqual(appPrice, "N/A");
				
			Report.Log(ReportLevel.Success, "Validation", "App Price including taxes is expected: " + appPriceTax);
			Validate.AreEqual(appPriceTax, "N/A");	
				
			//Close Browser
			//Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
						
		}
	}
}
