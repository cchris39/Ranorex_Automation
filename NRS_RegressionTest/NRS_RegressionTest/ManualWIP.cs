/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-11-16
 * Sceanrio: Lender create new manual entry
 * Time: 2:22 PM
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

namespace NRS_RegressionTest
{
	/// <summary>
	/// Description of ManualWIP.
	/// </summary>
	[TestModule("5E2743D8-8376-42B9-8760-30C88481B355", ModuleType.UserCode, 1)]
	public class ManualWIP : ITestModule
	{
		public static NRS_RegressionTestRepository repo = NRS_RegressionTestRepository.Instance;
		static ManualWIP instance = new ManualWIP();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ManualWIP()
		{
			// Do not delete - a parameterless constructor is required!
		}

		#region Varilables
		
		string _varLoan = "";
		[TestVariable("5819d994-eca7-4c32-9ab4-444c22b3410f")]
		public string varLoan
		{
			get { return _varLoan; }
			set { _varLoan = value; }
		}
		
		#endregion
		
		private string taskName = "Full Appraisal request";
		private string taskDesc = "Test only, please do not proceed, Thanks";
		private string asgValue = "0";
		
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		
		/// <summary>
		/// Create New Manual Entry
		/// </summary>
		public void newManualEntry(string task, string taskDate, string selValue, string taskDesc)
		{
			//Go to Manual WIP page
			repo.NRS.TopMenu.ManualWIP.Click();
			Delay.Milliseconds(300);
			
			//Create New Manual Entry
			repo.NRS.ManualWIP.TitleInput.PressKeys(task);
			Delay.Milliseconds(100);
			
			repo.NRS.ManualWIP.TaskDuedate.TagValue = taskDate;
			Delay.Milliseconds(100);
			
			repo.NRS.ManualWIP.AssignToSelect.TagValue = selValue;
			Delay.Milliseconds(100);
			
			repo.NRS.ManualWIP.DescriptionTextArea.PressKeys(taskDesc);
			Delay.Milliseconds(100);
			
			repo.NRS.ManualWIP.CreateWIPEntry.Click();
			Delay.Milliseconds(500);
			
			//Report Status
			Validate.Exists(repo.NRS.EmailSentSuccessfully);
			Report.Log(ReportLevel.Success, "Success", "New task: '" + taskName + "' created successfully.");
			
			
		}
		
		
		/// <summary>
		/// Get New Task name from 'Manual WIP' page
		/// </summary>
		static string getNewTaskName()
		{
			string taskText = repo.NRS.ManualWIP.New_TaskName.InnerText.Trim();
			int strLength = taskText.IndexOf(':') - taskText.IndexOf('-') -13;
			taskText = taskText.Substring(7,strLength).Trim();
			
			return taskText;
		}
		
		
		//Main
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			Delay.Milliseconds(100);
			
			string newDate = System.DateTime.Today.AddDays(10).ToString("yyyy-MM-dd");
						
			//Create New Manual Entry
			newManualEntry(taskName, newDate, asgValue, taskDesc);
			Delay.Milliseconds(200);
			
		}
	}
}
