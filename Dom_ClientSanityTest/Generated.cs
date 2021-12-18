using Microsoft.VisualBasic;
//
// Created by Ranorex
// User: an.zhou
// Date: 04/12/2015
// Time: 12:27 PM
//
// To change this template use Tools | Options | Coding | Edit Standard Headers.
//
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

namespace SanityTestSuite
{
	/// <summary>
	/// Description of ViewRequest.
	/// </summary>
	[TestModule("25317D1A-0A38-464D-AF06-B04C9358F2BF", ModuleType.UserCode, 1)]
	public class ViewRequest : ITestModule
	{
		static DomNasRepository repo = DomNasRepository.Instance;
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ViewRequest()
		{
			// Do not delete - a parameterless constructor is required!
		}

		public void KillBrowser(string IE)
		{
		}

		public void ClearBrowserCookies(string IE)
		{
		}
		#region "Variables"
		string _varUid = "";
		[TestVariable("65EC5536-6D5C-4ACA-B29E-DD7AF2720050")]
		public string varUid {
			get { return _varUid; }
			set { _varUid = value; }
		}

		string _varPwd = "";
		[TestVariable("763DDE07-7F02-4553-915E-9B95A9F12047")]
		public string varPwd {
			get { return _varPwd; }
			set { _varPwd = value; }
		}

		string _varNasNbr = "";
		[TestVariable("22FA0E19-502E-422B-8B33-05C8023852A1")]
		public string varNasNbr {
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}

		string _VarStrNbr = "";
		[TestVariable("AD09ED6D-3799-4CAC-8B23-8E97E9028FBD")]
		public string VarStrNbr {
			get { return _VarStrNbr; }
			set { _VarStrNbr = value; }
		}

		string _varPrice = "";
		[TestVariable("41076065-B047-4BFE-B11E-7B7DD8F02991")]
		public string varPrice {
			get { return _varPrice; }
			set { _varPrice = value; }
		}

		string _varAmount = "";
		[TestVariable("A56632E7-E93D-4F48-A8BE-11A479962366")]
		public string varAmount {
			get { return _varAmount; }
			set { _varAmount = value; }
		}

		string _varPostCode = "";
		[TestVariable("47E4CBAF-D2FA-468F-8261-8F90A74B348C")]
		public string varPostCode {
			get { return _varPostCode; }
			set { _varPostCode = value; }
		}

		#endregion

		string _varRef = "";
		[TestVariable("29746A1D-5866-46C5-954F-9C5623ECC31B")]
		public string varRef {
			get { return _varRef; }
			set { _varRef = value; }
		}

		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(Of ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		public void Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;

			Host.Local.OpenBrowser("http://test.nationwideappraisals.com", "IE", "", false, false, false, false, false);
			Delay.Milliseconds(100);

			Host.Local.ClearBrowserCookies("IE");
			Delay.Milliseconds(100);

			repo.DomesticNas.Userid.PressKeys(varUid);
			Delay.Milliseconds(0);

			repo.DomesticNas.Password.PressKeys(varPwd);
			Delay.Milliseconds(0);

			repo.DomesticNas.Submit.Click();
			Delay.Milliseconds(200);

			//Search Nas order
			repo.DomesticNas.SearchFilter.Click();
			repo.DomesticNas.MenuDisplay1.NasReqNum.PressKeys(varNasNbr);
			//varNasNbr
			repo.DomesticNas.MenuDisplay1.searchSubmit.Click();
			Delay.Milliseconds(100);

			repo.DomesticNas.MenuDisplay1.ViewRequestBtn.Click();
			Delay.Milliseconds(100);

			//Decalre varilable for view request page
			string viewReqNbr = repo.DomesticNas.MenuDisplay1.NASRequestNumber.InnerText;
			// Do not trim text
			viewReqNbr = viewReqNbr.Substring(25, 7);
			//msgbox(viewReqNbr)

			string SerType = repo.DomesticNas.MenuDisplay1.ViewServiceTypeTag.InnerText.Trim();
			//msgbox (SerType)

			string StrNbr = repo.DomesticNas.MenuDisplay1.TdTagPropertyAddress.InnerText.Trim();
			StrNbr = StrNbr.Substring(0, 3);
			//msgbox (StrNbr)

			string postCode = repo.DomesticNas.MenuDisplay1.TdPostCode.InnerText.Trim();
			postCode = postCode.Substring(0, 7);
			//msgbox (postCode)

			string refNbr = repo.DomesticNas.MenuDisplay1.TdTagClientRefNbr.InnerText.Trim();
			//msgbox (refNbr)

			string price = repo.DomesticNas.MenuDisplay1.TdTagPurchasePrice.InnerText.Trim();
			price = price.Remove(0, 1).Replace(".0", " ").Trim();
			//msgbox(price)

			string amount = repo.DomesticNas.MenuDisplay1.TdTagMortgageAmount.InnerText.Trim();
			amount = amount.Remove(0, 1).Replace(".0", " ").Trim();
			//msgbox(amount)

			//Report View Request status
			Report.Log(ReportLevel.Success, "Validation", viewReqNbr + " View Request Page; Nas Number is match.");
			Validate.AreEqual(viewReqNbr, varNasNbr);
			//varNasNbr
			Report.Log(ReportLevel.Info, "Validation", varNasNbr + " Service type is: " + SerType);
			Validate.Exists(repo.DomesticNas.MenuDisplay1.NASRequestNumber);
			Delay.SpeedFactor = 1.0;

			//Compare certain fields value with the input value from the ('Submit request') module
			Report.Log(ReportLevel.Success, "Validation", " Street Number is: " + StrNbr + "; Street Number match.");
			Validate.AreEqual(StrNbr, VarStrNbr);
			//varStrNbr

			Report.Log(ReportLevel.Success, "Validation", " Post Code is: " + postCode + "; Post Code match.");
			Validate.AreEqual(postCode, varPostCode);
			//varPostCode

			Report.Log(ReportLevel.Success, "Validation", " Purchase price is: " + price + "; Purchase Price match.");
			Validate.AreEqual(price, varPrice);
			//varPrice

			Report.Log(ReportLevel.Success, "Validation", " Mortgage is: " + amount + "; Mortgage amount match.");
			Validate.AreEqual(amount, varAmount);
			//varAmount

			Report.Log(ReportLevel.Success, "Validation", " Client Refence Number is: " + refNbr + "; CLient Reference number is match");
			Validate.AreEqual(refNbr, varRef);
			Delay.SpeedFactor = 1.0;


			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(0);
		}

	}
}
