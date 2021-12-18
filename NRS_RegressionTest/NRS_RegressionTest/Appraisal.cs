/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-11-17
 * Scenario: Lender Add Apprasal History by ordering a 'Drive-By' appraisal
 * Time: 10:46 AM
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

namespace NRS_RegressionTest
{
	/// <summary>
	/// Description of Appraisal.
	/// </summary>
	[TestModule("AAD85A38-0131-4946-B134-3E334F781529", ModuleType.UserCode, 1)]
	public class Appraisal : ITestModule
	{
		public static NRS_RegressionTestRepository repo = NRS_RegressionTestRepository.Instance;
		static Appraisal instance = new Appraisal();
				
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public Appraisal()
		{
			// Do not delete - a parameterless constructor is required!
		}

		#region Varilables
		
		string _varLoan = "";
		[TestVariable("bf2eca90-bda9-43c4-9f2f-b4ef19ddfe5c")]
		public string varLoan
		{
			get { return _varLoan; }
			set { _varLoan = value; }
		}
		
		#endregion
		
		private const string ID = "AP-";
		private string ordType = "DRIVE_BY";
		private string ordStatus = "ORDERED";
		private string apValue = "550000.0";
		
		/// <summary>
		/// Add Appraisal History
		/// </summary>
		public void addAppHistory(string id, string type, string status, string asValue)
		{
			//Goto To Appraisal Info page
			repo.NRS.TopMenu.Appraisal.MoveTo();
			Delay.Milliseconds(100);
			repo.NRS.Appraisal.AppraisalInformation.Click();
			Delay.Milliseconds(300);
			
			//Add Appraisal History
			repo.NRS.Appraisal.AddAppraisalHistory.Click();
			Delay.Milliseconds(600);
			
			//Fill Out Information
			repo.NRS.Appraisal.Appraisal_ID.PressKeys(id);
			Delay.Milliseconds(100);
			
			repo.NRS.Appraisal.Appraisal_Type.TagValue = type;
			Delay.Milliseconds(100);
			
			repo.NRS.Appraisal.Appraisal_Status.TagValue = status;
			Delay.Milliseconds(100);
			
			string ordDate = System.DateTime.Today.ToString("yyyy-MM-dd");
			string dueDate = System.DateTime.Today.AddDays(6).ToString("yyyy-MM-dd");
			repo.NRS.Appraisal.Appraisal_OrderDate.TagValue = ordDate;
			Delay.Milliseconds(100);
			repo.NRS.Appraisal.Appraisal_DueDate.TagValue =dueDate;
			Delay.Milliseconds(100);
			
			repo.NRS.Appraisal.Appraisal_AsIsValue.PressKeys(asValue);
			Delay.Milliseconds(100);
			
			repo.NRS.Appraisal.SaveBtn.Click();
			Delay.Milliseconds(300);
			
			//Verify Save Status
			Validate.Exists(repo.NRS.SuccessfullySavedRecordAppraisalHistory);
			Report.Log(ReportLevel.Success, "Success", "Appraisal History successfully added for: " + id);
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
			
			//Add Appraisal History
			var random = new Random();
			var idNbr = random.Next(1, 999999);
			string apID = ID + idNbr.ToString();
			
			addAppHistory(apID, ordType, ordStatus, apValue);
			Delay.Milliseconds(100);
			
		}
	}
}
