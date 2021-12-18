/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 25/01/2016
 * Time: 3:27 PM
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

namespace Dom_AppraiserSanityTest
{
	[TestModule("41C516DC-48B4-40E0-BC05-0C1A857A58BC", ModuleType.UserCode, 1)]
	public class AppraiserAcceptRequest : ITestModule
	{
		/// <summary>
		/// Using the Dom_AppraiserSanityTestRepository repository.
		/// </summary>
		public static Dom_AppraiserSanityTestRepository repo = Dom_AppraiserSanityTestRepository.Instance;
		
		static AppraiserAcceptRequest instance = new AppraiserAcceptRequest();
		/// <summary>
	/// <summary>
	/// Description of AppraiserAcceptRequest.
	/// Scenario: Appraiser accept client submit request; This module can be exclude as the 'Accept' action also include in the 'Appraisercomment' module
	/// </summary>
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public AppraiserAcceptRequest()
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
		[TestVariable("5376E41D-76C4-43F7-8EFD-BBF9D9C54A38")]
		public string varAppraiser
		{
			get { return _varAppraiser; }
			set { _varAppraiser = value; }
		}
		
		string _varPwd = "";
		[TestVariable("8DB15536-254D-43EF-8F82-20921A659CC4")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("064A7AE3-4A43-499C-8745-3A9A4A0F6FB4")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varUrl = "";
		[TestVariable("CF26F2AA-9080-4B71-853F-0A6147886BF7")]
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
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			Host.Local.ClearBrowserCookies("IE");
			Delay.Milliseconds(100);
			
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(100);
			
			//WebDocument DomWebDocu = "/dom[@domain='uattest.nas.com']";
			
			//Appraiser log in
			repo.DomNasHome.Userid.PressKeys(varAppraiser);
			repo.DomNasHome.Password.PressKeys(varPwd);
			repo.DomNasHome.Submit.Click();
			Delay.Milliseconds(100);
			
			//Search By Nas Number
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys("varNasNbr");     // varNasNbr
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			
			//Accepted Request
			string chgStatus = "Accepted";
			
			repo.DomNasHome.MenuDisplay.Chastatus.Element.SetAttributeValue("TagValue", chgStatus);
			repo.DomNasHome.MenuDisplay.ChgStatusBtn.Click();
			Delay.Milliseconds(100);
			
			Report.Log(ReportLevel.Info, "Validation", "Status changed for: " + "varNasNbr");
			Validate.Exists(repo.DomNasHome.MenuDisplay.StatusChangedFromNew);
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(0);
			
		}
	}
}
