/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 14/03/2016
 * Time: 4:10 PM
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
	/// Description of ViewAppraisalHistory.
	/// </summary>
	[TestModule("85DF4112-FA0A-453E-B902-95BE495DD3C9", ModuleType.UserCode, 1)]
	public class ViewAppraisalHistory : ITestModule
	{
		#region Varilables
		string _varNasNbr = "";
		[TestVariable("DDF2BE9D-C8B0-4C9B-8FE0-B9C83B9DEC0D")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		#endregion
		
		/// <summary>
		/// Using the Dom_SanityTestRepository repository.
		/// </summary>
		public static BrokerFlowRepository repo = BrokerFlowRepository.Instance;
		
		static ViewAppraisalHistory instance = new ViewAppraisalHistory();
				
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ViewAppraisalHistory()
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
			
			Delay.Milliseconds(100);
			//Search By Nas Number
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(varNasNbr);       //varNasNbr
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			repo.DomNasHome.MenuDisplay.AppraisalHistoryBtn.Click();
			
			string status = repo.DomNasHome.MenuDisplay.TdTagStatus.InnerText.Trim();
			string appPrice = repo.DomNasHome.MenuDisplay.TdTagAppPrice.InnerText.Trim();
			string appPriceTax = repo.DomNasHome.MenuDisplay.TdTagAppPriceTax.InnerText.Trim();
			string nasNbr = repo.DomNasHome.MenuDisplay.AppraisalRequestNumber.InnerText.Trim().Substring(26,7);
			string regionType =repo.DomNasHome.MenuDisplay.RegionType.InnerText.Trim().Substring(13,5);
			
					
			//Report Appraisal History Status and validate Status, App Price etc.,
			Report.Log(ReportLevel.Info, "Validation", "Appraisal History page present.");
			Validate.Exists(repo.DomNasHome.MenuDisplay.AppraisalRequestHistoryReport);
			Report.Log(ReportLevel.Success, "Validation", "Nas number: " + varNasNbr  + " match");
			Validate.AreEqual(nasNbr, varNasNbr);
			Report.Log(ReportLevel.Info, "Validation", "Region Type is: " + regionType);
			
			Report.Log(ReportLevel.Success, "Validation", "Request status is expected: " + status);
			Validate.AreEqual(status, "Credit Pending");
				
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
