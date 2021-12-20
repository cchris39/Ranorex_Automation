/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 12/02/2016
 * Time: 12:09 PM
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

namespace NasAdmin
{
	/// <summary>
	/// Description of FeeIncrease.
	/// </summary>
	[TestModule("D76F5CCE-61E3-4556-AB80-3864299FFD14", ModuleType.UserCode, 1)]
	public class FeeIncrease : ITestModule
	{
		/// <summary>
		/// Using the NasAdminRepository repository.
		/// </summary>
		public static NasAdminRepository repo = NasAdminRepository.Instance;
		
		static FeeIncrease instance = new FeeIncrease();
		
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("6CE10578-A611-4F70-BEF7-4A7990484093")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varAdmin = "";
		[TestVariable("3E0812A7-AF76-4992-8782-9318356E0FE0")]
		public string varAdmin
		{
			get { return _varAdmin; }
			set { _varAdmin = value; }
		}
		
		string _varPwd = "";
		[TestVariable("B21B455C-3BFA-45CD-A719-37C86355CF76")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varUrl = "";
		[TestVariable("DDC91614-6AD4-4E7C-A36B-61EDAF92BE51")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		#endregion
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public FeeIncrease()
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
			
			
			
			repo.Dom_NasHome.NasAdminMenu.SearchFilter.Click();
			Delay.Milliseconds(100);	
			
			//Input Nas Number to search
			repo.Dom_NasHome.MenuDisplay.NasReqNumSearch.PressKeys(varNasNbr);
			repo.Dom_NasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			//Check request status
			string status = repo.Dom_NasHome.MenuDisplay.ReqStatus.InnerText.Trim();
			
			if (status == "New" | status == "Accepted" | status == "Appointment")
			{
				repo.Dom_NasHome.NasAdminFunction.RequestFeeIncrease.Click();
				Delay.Milliseconds(100);
			}else
			{
				Report.Warn ("Request " + varNasNbr + " is in " + status + "; No fee request can be submitted.");
			}
			
			string[] fee1 = {"Executive Property", "1.0M - 1.5M", "75.00"};
			string[] fee2 = {"Other (please specify)", "Test other fee", "45.00"};
			string feeNote = "Test fee increase.";
			
			//Submit Fee reason
			repo.Dom_NasHome.MenuDisplay.Fee.AddiFee1Desc.TagValue = fee1[0];
			repo.Dom_NasHome.MenuDisplay.Fee.Renovation1.TagValue = fee1[1];
			Delay.Milliseconds(100);
			
			repo.Dom_NasHome.MenuDisplay.Fee.AddiFee1.Click();
			repo.Dom_NasHome.MenuDisplay.Fee.AddiFee1.PressKeys(fee1[2]);
			Delay.Milliseconds(100);
			
			repo.Dom_NasHome.MenuDisplay.Fee.AddiFee2Desc.Click();
			repo.Dom_NasHome.MenuDisplay.Fee.AddiFee2Desc.TagValue = fee2[0];
			repo.Dom_NasHome.MenuDisplay.Fee.OtherDesc2.PressKeys(fee2[1]);
			Delay.Milliseconds(100);
			
			repo.Dom_NasHome.MenuDisplay.Fee.AddiFee2.Click();
			repo.Dom_NasHome.MenuDisplay.Fee.AddiFee2.PressKeys(fee2[2]);
			Delay.Milliseconds(100);
			
			repo.Dom_NasHome.MenuDisplay.Fee.FeeNotes.PressKeys(feeNote);
			
			//Get new total fee
			string newTotalFee = repo.Dom_NasHome.MenuDisplay.Fee.TotalNewFee.InnerText.Trim();
			
			//Submit Fee Increase request
			repo.Dom_NasHome.MenuDisplay.Fee.FeeSubmitButton.Click();
			Delay.Milliseconds(300);
			
			//Report Fee Submit Status
			Report.Log(ReportLevel.Success, "Validation", "Fee increase request successfully submitted for: " + varNasNbr );
			Report.Log(ReportLevel.Info, "Information", "New Total Fee for: " + varNasNbr + " is " + newTotalFee);
			Validate.Exists(repo.Dom_NasHome.MenuDisplay.Fee.FeeIncreaseRecordCreatedSuccessfully);
			Delay.Milliseconds(200);
			
			//Verify Fee information in Fee adminstration and update Fee Status
			repo.Dom_NasHome.NasAdminMenu.SystemAdmin.Click();
			repo.Dom_NasHome.NasAdminMenu.FeeAdministration.Click();
			Delay.Milliseconds(200);	
					
			repo.Dom_NasHome.MenuDisplay.Fee.Appreqnum.PressKeys(varNasNbr);
			repo.Dom_NasHome.MenuDisplay.Fee.FeeSearchBtn.Click();
			Delay.Milliseconds(100);
			
			repo.Dom_NasHome.MenuDisplay.Fee.FeeAdmin_Radio.Click();
			repo.Dom_NasHome.MenuDisplay.Fee.ViewFeeDetailsBtn.Click();
			Delay.Milliseconds(100);
			
			string newFee = repo.Dom_NasHome.MenuDisplay.Fee.FeeAdminAppraiserNewFee.InnerText.Trim();
			string newPrice = repo.Dom_NasHome.MenuDisplay.Fee.FeeAdminClientNewPrice.InnerText.Trim();
			
			
			fee1[2] = "$" + fee1[2];
			fee2[2] = "$" + fee2[2];
			
			string feeAmount1 = repo.Dom_NasHome.MenuDisplay.Fee.FeeAdminAddiFee1.InnerText.Trim();
			string feeReason1 = repo.Dom_NasHome.MenuDisplay.Fee.FeeAdminAddiFee1Reason.InnerText.Trim();
			string feeDesc1 = repo.Dom_NasHome.MenuDisplay.Fee.FeeAdminAddiFee1Desc.InnerText.Trim();
			
			string feeAmount2 = repo.Dom_NasHome.MenuDisplay.Fee.FeeAdminAddiFee2.InnerText.Trim();
			string feeReason2 = repo.Dom_NasHome.MenuDisplay.Fee.FeeAdminAddiFee2Reason.InnerText.Trim();
			string feeDesc2 = repo.Dom_NasHome.MenuDisplay.Fee.FeeAdminAddiFee2Desc.InnerText.Trim();
			
			string[] feeInfo1 = {feeReason1,feeDesc1,feeAmount1};
			string[] feeInfo2 = {feeReason2,feeDesc2,feeAmount2};
			
			//Compare Two Fee String Array
			bool assert1 = ArraysAreEqual(fee1, feeInfo1);
			bool assert2 = ArraysAreEqual(fee2, feeInfo2);
			
			//Report Fee status
			if (assert1)
			{
				Report.Log(ReportLevel.Success, "Validation", "Fee increase record 1 for appraiser match." );
				Report.Log(ReportLevel.Info, "Information", "Fee1 amount is: " + feeAmount1 + "; " + "Fee1 reason is: " + feeReason1 + "; " + "Fee1 note is: " + feeDesc1 + ". ");
				
			}else
			{
				Report.Log(ReportLevel.Failure, "Validation", "Fee increase record 1 for appraiser match." );
				Report.Log(ReportLevel.Info, "Information", "Fee1 amount is: " + feeAmount1 + "; " + "Fee1 reason is: " + feeReason1 + "; " + "Fee1 note is: " + feeDesc1 + ". ");
			}
			
			if (assert2)
			{
				Report.Log(ReportLevel.Success, "Validation", "Fee increase record 2 for appraiser match." );
				Report.Log(ReportLevel.Info, "Information", "Fee2 amount is: " + feeAmount2 + "; " + "Fee2 reason is: " + feeReason2 + "; " + "Fee2 note is: " + feeDesc2 + ". ");
				
			}else
			{
				Report.Log(ReportLevel.Failure, "Validation", "Fee increase record 2 for appraiser match." );
				Report.Log(ReportLevel.Info, "Information", "Fee2 amount is: " + feeAmount2 + "; " + "Fee2 reason is: " + feeReason2 + "; " + "Fee2 note is: " + feeDesc2 + ". ");
			}
			
			//Sys admin approved fee
			Delay.Milliseconds(100);
			repo.Dom_NasHome.MenuDisplay.Fee.FeeStatusOpt.TagValue = "Approved";
			Delay.Milliseconds(100);
			repo.Dom_NasHome.MenuDisplay.Fee.UpdateFeeButton.Click();
			Delay.Milliseconds(200);
			
			
			
			//Report Fee update status
			Report.Log(ReportLevel.Success, "Validation", "Fee increase has been updated successfully for " + varNasNbr);
			Validate.Exists(repo.Dom_NasHome.MenuDisplay.Fee.H2TagFeeIncreaseHasBeenUpdatedSuccess);
			Report.Log(ReportLevel.Info, "Information", "New total appraiser fee is: " + newFee + "; ");
			Report.Log(ReportLevel.Info, "Information", "New total client price is: " + newPrice + ". ");
			
			//Close Browser
			//Host.Local.KillBrowser("IE");
			Delay.Milliseconds(300);
				
		}
		
		    static bool ArraysAreEqual(string[] x, string[] y) 
		    {
            if (x.Length != y.Length) {
                return false;
            }

            for (int index = 0; index < x.Length; index++)
            {
                if (x[index] != y[index]) {
                    return false;
                }
            }

            return true;
        	}
	}
}
