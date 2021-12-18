/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-11-17
 * Time: 12:25 PM
 * Scenario: Lender add new claim as 'TITLE_TRANSFER'
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

namespace NRS_RegressionTest
{
	/// <summary>
	/// Description of Insurer.
	/// </summary>
	[TestModule("BB1EE832-DA7E-4BF7-86C8-D414EE0A6050", ModuleType.UserCode, 1)]
	public class Insurer : ITestModule
	{
		public static NRS_RegressionTestRepository repo = NRS_RegressionTestRepository.Instance;
		static Insurer instance = new Insurer();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public Insurer()
		{
			// Do not delete - a parameterless constructor is required!
		}

		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		
		
		#region Varilables
		
		string _varLoan = "";
		[TestVariable("34682a58-8bb6-4581-bd64-fdcece42067a")]
		public string varLoan
		{
			get { return _varLoan; }
			set { _varLoan = value; }
		}
		
		#endregion
		
		private string amount = "630000.0";
		private string claimType = "TITLE_TRANSFER";
		
		/// <summary>
		/// Add New Main Claim (Claim Type: 'TITLE_TRANSFER')
		/// </summary>
		public void addMainClaim(string type, string date, string amt)
		{
			//Add New Main Claim
			repo.NRS.Insurer.AddMainClaim.Click();
			Delay.Milliseconds(300);
			
			repo.NRS.Insurer.ClaimType.TagValue = type;
			Delay.Milliseconds(100);
			repo.NRS.Insurer.ClaimFiledInputDate.PressKeys(date);
			Delay.Milliseconds(100);
			repo.NRS.Insurer.ClaimAmount.PressKeys(amt);
			Delay.Milliseconds(100);
			
			repo.NRS.Insurer.SaveClaim.Click();
			Delay.Milliseconds(300);
			
			//Report Status
			Validate.Exists(repo.NRS.Record_SuccessfullySavedBox);
			Report.Log(ReportLevel.Success, "Success", "New Main Claim created and saved. Claim type: "+ type + "; Claim Filed: " + date + "; Claim Amount: " + amt);
			
		}
		
		
		/// <summary>
		/// Add New Main Claim (Claim Type: 'TITLE_TRANSFER')
		/// </summary>
		public void addSuplClaim(string type, string date, string amt)
		{
			//Add New Supplement Claim
			repo.NRS.Insurer.AddSupplementClaim.Click();
			Delay.Milliseconds(300);
			
			repo.NRS.Insurer.SupplementClaimType.TagValue = type;
			Delay.Milliseconds(100);
			repo.NRS.Insurer.SupplementClaimFileDate.PressKeys(date);
			Delay.Milliseconds(100);
			repo.NRS.Insurer.SupplementClaimAmount.PressKeys(amt);
			Delay.Milliseconds(100);
			
			repo.NRS.Insurer.SupplementClaimSaveBtn.Click();
			Delay.Milliseconds(300);
			
			//Report Status
			Validate.Exists(repo.NRS.Record_SuccessfullySavedBox);
			Report.Log(ReportLevel.Success, "Success", "New Supplement Claim created and saved. Claim type: "+ type + "; Claim Filed: " + date + "; Claim Amount: " + amt);
		}
		
		
		//Main
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			Delay.Milliseconds(100);
			string toDate = System.DateTime.Today.ToString("yyyy-MM-dd");
			
			//Goto Mortgage Insurer page
			repo.NRS.TopMenu.MortgageInsurer.MoveTo();
			Delay.Milliseconds(100);
			repo.NRS.Insurer.MortgageInsurer.Click();
			Delay.Milliseconds(200);
			
			//Check if Main Claim exist -If 'Yes' add Supplement Claim, else add New Main Claim; Claim type: 'TITLE_TRANSFER'.
			if(repo.NRS.Insurer.AddSupplementClaimInfo.Exists(200))
			{
				addSuplClaim(claimType, toDate, amount);
			}else
			{
				addMainClaim(claimType, toDate, amount);
			}
			Delay.Milliseconds(200);
			
		}
	}
}
