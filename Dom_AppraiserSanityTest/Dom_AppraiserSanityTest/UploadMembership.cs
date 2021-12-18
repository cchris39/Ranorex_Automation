/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 27/01/2016
 * Time: 11:30 AM
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
using System.Data;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace Dom_AppraiserSanityTest
{
	/// <summary>
	/// Description of AppraiserRepository.
	/// </summary>
	[TestModule("A4513EA3-FF29-4264-B2D5-4CDC390FA554", ModuleType.UserCode, 1)]
	public class UploadMembership : ITestModule
	{
		/// <summary>
		/// Using the Dom_AppraiserSanityTestRepository repository.
		/// </summary>
		public static Dom_AppraiserSanityTestRepository repo = Dom_AppraiserSanityTestRepository.Instance;
		
		static UploadMembership instance = new UploadMembership();
		
		
		#region Varilables
		
		string _varUrl = "";
		[TestVariable("BB82EC8D-1A66-47DA-809B-1287DCE0A07E")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		string _varPwd = "";
		[TestVariable("A00DBC9D-BBBF-4339-AF81-A04D69F9EB4A")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varAppraiser = "";
		[TestVariable("9E684281-EFB9-4AED-A471-AAEA5607E28F")]
		public string varAppraiser
		{
			get { return _varAppraiser; }
			set { _varAppraiser = value; }
		}
		
		#endregion
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public UploadMembership()
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
			
			Host.Local.ClearBrowserCookies("IE");
			Delay.Milliseconds(200);
			
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(300);
			
			WebDocument DomWebDocu = repo.DomNasHome.Self;
			
			//Appraiser log in
			repo.DomNasHome.Userid.PressKeys(varAppraiser);   
			repo.DomNasHome.Password.PressKeys(varPwd);
			repo.DomNasHome.Submit.Click();
			Delay.Milliseconds(200);
			
			repo.SavePwdNo.Click();
			
			repo.DomNasHome.UploadViewEOAndMembership.Click();
				
			const string varFstName = "FstName";
			const string varLstname = "LstName";
						
			const string missDateWarning = "Expiry date is required to upload an E&O document.";
			const string dateWarning = "Expiry date must be greater than current date.";
			const string licenWarning = "Designation is a mandatory field for proof of designation documents.";
			const string memberWarning = "Member/Proof of Designation Number is a mandatory field for Proof of Membership.";
			const string otherWarning = "Please specify type of document when you pick file type as Other.";
			const string missFileWarning = "Please choose a file to upload.";
			const string fileTypeWarning = "File type is a mandatory field.";
			const string provWarning = "Province is a mandatory field.";
			
			
			
			//const string dupFileMsgWarning = "For file:house1.jpg, there is another document with the same information. Do you wish to replace this file? Click 'Ok' to replace, click 'Cancel' to not to replace.";
			const string otherTxtDesc = "ORDER of AACI - Award winning appraiser";
			const string otherTxtDesc2 = "Other test document";
			const string licNbr = "WMD-215354";
			
			var curDate = System.DateTime.Today;
			var invalidExpiryDate = curDate.AddDays(-90).ToString("yyyy-MM-dd");
			var expiryDate = curDate.AddDays(90).ToString("yyyy-MM-dd");
						
			string [] warning = new string[8] {missDateWarning, dateWarning, licenWarning, memberWarning, otherWarning, missFileWarning, fileTypeWarning, provWarning};
			
			string fileLoad = @"c:\Upload_Files\test.pdf"; 
			
			
			//Test Upload Document section (Pending Replace file coding)
			for (int i = 0; i <= 3; i++)
       		{
				string [] msg = new string [8];
				
				int j = i + 1;
				int x = j + 1;
				int y = x + 2;
				//Define xpath for 'Upload Document' section
				string usrNamePath = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[@name ='appraiser$']".Replace("$", j.ToString());
				string fstNamePath = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[@name = 'firstname$']".Replace("$", j.ToString());
				string lstNamePath = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[@name = 'lastname$']".Replace("$", j.ToString());
				string docuPath = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[@name = 'document$']".Replace("$", j.ToString());
				string fileTypePath = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//select[@name = 'filetype$']".Replace("$",j.ToString());
				string provPath = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//select[@name = 'state$']".Replace("$", j.ToString());
				string desgPath = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//select[@name = 'designation$']".Replace("$",j.ToString());
				string licenPath = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[@name = 'license$']".Replace("$", j.ToString());
				string expDatePath = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[@name = 'expiry_date$']".Replace("$", j.ToString());
				string otherTxtPath = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//input[@name = 'filetypeother$' and @formaction=null()]".Replace("$", j.ToString());
				string dateEraserPath = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//form[#'uploadlicense']/table//tr[$]//img[@alt='Click here to clear the date field.']".Replace("$", x.ToString());
				
				Ranorex.InputTag usrName = usrNamePath;
				Ranorex.InputTag fstName = fstNamePath;
				Ranorex.InputTag lstName = lstNamePath;
				Ranorex.InputTag docuUpload = docuPath;
				Ranorex.SelectTag fileType = fileTypePath;
				Ranorex.SelectTag province = provPath;
				Ranorex.SelectTag designation = desgPath;
				Ranorex.InputTag license = licenPath;
				Ranorex.InputTag expDate = expDatePath;
				Ranorex.InputTag otherTxt = otherTxtPath;
				Ranorex.ImgTag dateErase = dateEraserPath;
				
				
				//Processing each row
				//Fill out username, first and last name, province, etc.,
				
				
					//Keyboard.Press("{Tab}");
					if (i!=3)
					{
					usrName.PressKeys(varAppraiser);
					Keyboard.Press("{Tab}");
					fstName.PressKeys(varFstName + j.ToString());
					Keyboard.Press("{Tab}");
					lstName.PressKeys(varLstname + j.ToString());
					Keyboard.Press("{Tab}");
					province.TagValue = "NS";	
						
					docuUpload.DoubleClick();
					Delay.Milliseconds(100);
					repo.ChooseFileToUpload.FileName.PressKeys(fileLoad);
					Keyboard.Press("{Return}");
					Delay.Milliseconds(300);
					}
				
					//Processing each case
					switch (i) 
					{
									case 0: // (Row1)Upload E&O and validate Missing and Invalid expiry date message 
											
											fileType.TagValue = "E&O";
												Keyboard.Press("{Tab}");
											designation.TagValue = "CRA";
											repo.DomNasHome.UploadMembershipEO.DocuUploadBtn.Click();
											Delay.Milliseconds(100);
											
											Ranorex.Form msgForm1 = "/form[@title='Message from webpage']";
											msg[i] = repo.MessageFromWebpage.TextWarning.TextValue.Trim();
											//Validate Missing Expiry Date warning message
											if (msgForm1 != null)
												{
													Report.Log(ReportLevel.Success, "Validation", "Missing expiry date warning message verified.");
													Validate.AreEqual(msg[i], warning[i]);
													repo.MessageFromWebpage.ButtonOK.Click();
												}else
												{
													Report.Log(ReportLevel.Failure, "Validation", "No missing expiry date warning message present, please check.");
													break;
												}	
													
											//Validate Invalid expiry date warning message
											Delay.Milliseconds(200);
											expDate.Element.SetAttributeValue("TagValue", invalidExpiryDate);
											repo.DomNasHome.UploadMembershipEO.DocuUploadBtn.Click();
													
											Ranorex.Form msgForm2 = "/form[@title='Message from webpage']";
											msg[j] = repo.MessageFromWebpage.TextWarning.Caption.Trim();
											
											if (msgForm2 != null)
												{	Report.Log(ReportLevel.Success, "Validation", "Invalid expiry date warning message verified.");
													Validate.AreEqual(msg[j], warning[j]);
													repo.MessageFromWebpage.ButtonOK.Click();
												
													//Erase invalid date and input valid Expiry date
													dateErase.Click();
													expDate.Element.SetAttributeValue("TagValue", expiryDate);
													Delay.Milliseconds(100);
													break;
												}else
												{
													Report.Log(ReportLevel.Failure, "Validation", "No invalid expiry date warning message present, please check.");
													break;
												}
											
											
									case 1: 	// (Row2)Upload Proof of Designation and validate missing designation Number message
									
											fileType.TagValue = "Proof of Designation";
												Keyboard.Press("{Tab}");
											repo.DomNasHome.UploadMembershipEO.DocuUploadBtn.Click();
											Delay.Milliseconds(100);
											
											Ranorex.Form msgForm3 = "/form[@title='Message from webpage']";
											msg[j] = repo.MessageFromWebpage.TextWarning.Caption.Trim();
											//Validate Missing Designation warning message
											if (msgForm3 != null)
												{
													Report.Log(ReportLevel.Success, "Validation", "Designation is a mandatory field for proof of designation documents.");
													Validate.AreEqual(msg[j], warning[j]);
													repo.MessageFromWebpage.ButtonOK.Click();
													
													//Input Designation type
													designation.TagValue = "CRA";
													repo.DomNasHome.UploadMembershipEO.DocuUploadBtn.Click();
													Delay.Milliseconds(100);
												}else
												{
													Report.Log(ReportLevel.Failure, "Validation", "No missing Designation warning message present, please check.");
													break;
												}
													
											//Validate Missing License Number warning message
											Ranorex.Form msgForm4 = "/form[@title='Message from webpage']";
											msg[x] = repo.MessageFromWebpage.TextWarning.Caption.Trim();		
											if (msgForm4 != null)
												{
													Report.Log(ReportLevel.Success, "Validation", "Member/Proof of Designation Number is a mandatory field.");
													Validate.AreEqual(msg[x], warning[x]);
													repo.MessageFromWebpage.ButtonOK.Click();
													Delay.Milliseconds(100);
													
													//Input missing License Number
													license.PressKeys(licNbr);
													Delay.Milliseconds(100);
													break;
																					
												}else
												{
													Report.Log(ReportLevel.Failure, "Validation", "No missing Member/Proof of Designation Number warning message present, please check.");
													break;
												}
																			
											
									case 2:   // (Row3)Upload Other document and validate blank other description field and missing choose file message
											fileType.TagValue = "Other";
											repo.DomNasHome.UploadMembershipEO.DocuUploadBtn.Click();
											Delay.Milliseconds(100);
											
											Ranorex.Form msgForm5 = "/form[@title='Message from webpage']";
											msg[x] = repo.MessageFromWebpage.TextWarning.Caption.Trim();
											if (msgForm5 != null)
												{
													Report.Log(ReportLevel.Success, "Validation", "Missing other description warning message verified.");
													Validate.AreEqual(msg[x], warning[x]);
													repo.MessageFromWebpage.ButtonOK.Click();
												
													//Input other text message
													otherTxt.PressKeys(otherTxtDesc);
													Delay.Milliseconds(100);
													break;
												}else
												{
													Report.Log(ReportLevel.Failure, "Validation", "No missing other description warning message present, please check.");
													break;
												}
												
									case 3:   // (Row4)Upload Other document and validate Not choose file and pick Province
											
											usrName.PressKeys(varAppraiser);
											Keyboard.Press("{Tab}");
											fstName.PressKeys("FStName" + j.ToString());
											Keyboard.Press("{Tab}");
											lstName.PressKeys("LStName" + j.ToString());
											Keyboard.Press("{Tab}");
											
											fileType.TagValue = "Other";
											Keyboard.Press("{Tab}");
											otherTxt.PressKeys(otherTxtDesc2);
											repo.DomNasHome.UploadMembershipEO.DocuUploadBtn.Click();
											Delay.Milliseconds(100);
											
											// Validate Not choose file to upload warning message
											Ranorex.Form msgForm6 = "/form[@title='Message from webpage']";
											msg[x] = repo.MessageFromWebpage.TextWarning.Caption.Trim();
											if (msgForm6 != null)
												{
													Report.Log(ReportLevel.Success, "Validation", "Not choose file warning message verified.");
													Validate.AreEqual(msg[x], warning[x]);
													repo.MessageFromWebpage.ButtonOK.Click();
												
													//Choose file to upload
													docuUpload.DoubleClick();
													Delay.Milliseconds(100);
													repo.ChooseFileToUpload.FileName.PressKeys(fileLoad);
													Keyboard.Press("{Return}");
													Delay.Milliseconds(300);
													
													repo.DomNasHome.UploadMembershipEO.DocuUploadBtn.Click();
													Delay.Milliseconds(100);
												}else
												{
													Report.Log(ReportLevel.Failure, "Validation", "No not choose file warning message present, please check.");
													break;
												}	
											
											//Validate missing Province warning message
											Ranorex.Form msgForm7 = "/form[@title='Message from webpage']";
											msg[y] = repo.MessageFromWebpage.TextWarning.Caption.Trim();
											if (msgForm7 != null)
												{
													Report.Log(ReportLevel.Success, "Validation", "Missing Province warning message verified.");
													Validate.AreEqual(msg[y], warning[y]);
													repo.MessageFromWebpage.ButtonOK.Click();
												
													//Pick Province
													province.TagValue = "NS";
													repo.DomNasHome.UploadMembershipEO.DocuUploadBtn.Click();
													Delay.Milliseconds(100);
													break;
												}else
												{
													Report.Log(ReportLevel.Failure, "Validation", "No Missing Province warning message present, please check.");
													break;
												}	
									
					}//End of switch case
			}//End of For loop
			
			//Upload attached file all at once (Total of Four files)
			Delay.Milliseconds(100);
			repo.DomNasHome.UploadMembershipEO.DocuUploadBtn.Click();
			Delay.Milliseconds(500);
			
			//Report Document upload status
			Report.Log(ReportLevel.Info, "Information", "File upload and saved successfully.");
			//Validate.Exists(repo.DomNasHome.UploadMembershipEO.FileUploadedSuccessful);
			
			
			//===============Delete an upload file=====================
			const string missCommentWarning = "A comment is required to delete a document. Comment should have at least 15 characters and be less than 300 characters.";
			
			//Ranorex.InputTag fileCheck = DomWebDocu.FindSingle(".//frameset[#'mainFrameset']/frame[@name='menu_display']//input[@tagname='input' and @id='checkbox0']");
			//fileCheck.Checked = "true";
			
			repo.DomNasHome.UploadMembershipEO.DocuCheckbox0.Click();
			Delay.Milliseconds(100);
			repo.DomNasHome.UploadMembershipEO.Deletebutton.Click();
			Delay.Milliseconds(200);
						
			//Validate missing remove document comment message
			Ranorex.Form msgForm8 = "/form[@title='Message from webpage']";
			string msg_removeComment = repo.MessageFromWebpage.TextWarning.Caption.Trim();
			
							if (msgForm8 != null)
												{
													Report.Log(ReportLevel.Success, "Validation", "No document remove reason warning message verified.");
													Validate.AreEqual(msg_removeComment, missCommentWarning);
													repo.MessageFromWebpage.ButtonOK.Click();
													Delay.Milliseconds(100);
													
													//Input remove reason
													repo.DomNasHome.UploadMembershipEO.DocuRemoveComment.PressKeys("Test remove document comment reason.");
													repo.DomNasHome.UploadMembershipEO.Deletebutton.Click();
													Delay.Milliseconds(200);
													repo.MessageFromWebpage.ButtonOK.Click();
													Delay.Milliseconds(100);
			
												}else
												{
													Report.Log(ReportLevel.Failure, "Validation", "No missing document remove reason warning message present, please check.");
												}	
			
			//Report delete file status
			Report.Log(ReportLevel.Success, "Validation", "Document has been removed successfully.");
			Validate.Exists(repo.DomNasHome.UploadMembershipEO.FileDeletedForFileConfirmed);
			Delay.Milliseconds(100);
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
			
		}
	}
}
