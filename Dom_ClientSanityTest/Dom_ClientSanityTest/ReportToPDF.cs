using Ranorex;
using Ranorex.Core.Reporting;
using Ranorex.Core.Testing;
using System;
using System.Net.Mail;
using System.Net;
using System.IO;


namespace ReportToPDF
{
	/// <summary>
	/// Allows conversion of Ranorex report files to PDF
	/// </summary>
	[TestModule("FFA0759D-37D2-4ABB-89A7-411F0FCF2DFE", ModuleType.UserCode, 1)]
	public class ReportToPDF : ITestModule
	{
		//Variables
		string PDFName;
		string xml;
		string details;
		static bool registered = false;
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ReportToPDF()
		{
			//Init variables
			this.PDFName = "";
			
			this.xml =  "";
			
			//Possible values: none | failed | all
			this.details = "all";
		}
		
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		void ITestModule.Run()
		{
			//Delegate must be registered only once
			if(!registered)
			{
				//PDF will be generated at the very end of the testsuite
				TestSuite.TestSuiteCompleted += delegate {
					
					//Specify report name if not already set
					if(String.IsNullOrEmpty(this.PDFName))
					{
						this.PDFName = CreatePDFName();
					}
					
					//Necessary to end the testreport in order to update the duration
					TestReport.EndTestModule();
					
					//Comment out if ConvertReportToPDF() is called directly
					try
					{
						Report.LogHtml(ReportLevel.Success,"PDFReport", "Successfully created PDF: <a href='" + ConvertReportToPDF(PDFName, xml, details) + "' target='_blank'>Open PDF</a>");
						//Add this send email method here ==================================================
						//Call the SendMail method after the PDF has been created.
                  		SendMail();

					}
					catch (Exception e)
					{
						Console.BackgroundColor = ConsoleColor.Black;
						Console.ForegroundColor = ConsoleColor.Red;
						Console.WriteLine("ReportToPDF: " + e.Message);
						Console.ResetColor();
						Console.WriteLine("Press any key to continue...");
						Console.ReadKey();
					}
					
					//Delete *.rxzlog if not enabled within test suite settings
					Cleanup();
					
					//Update error value
					UpdateError();
				};
				
				registered = true;
			}
		}
		
		public static string ConvertReportToPDF (string PDFName, string xml, string details)
		{
			var reportFileDirectory = TestReport.ReportEnvironment.ReportFileDirectory;
			var name = TestReport.ReportEnvironment.ReportName;
			
			var input = TestReport.ReportEnvironment.ReportZipFilePath;
			var PDFReportFilePath = String.Format(@"{0}\{1}", reportFileDirectory, CheckExtension(PDFName));
			
			TestReport.SaveReport();
			Report.Zip(TestReport.ReportEnvironment, reportFileDirectory, name);

			Ranorex.PDF.Creator.CreatePDF(input, PDFReportFilePath, xml, details);
			return PDFReportFilePath;
		}

		private static string CheckExtension(string PDFName)
		{
			var split = PDFName.Split('.');
			
			for(int i =0; i <split.Length; i++)
			{
				if(split[i].Contains("pdf") && i == split.Length -1)
				{
					return PDFName;
				}
			}
			
			return PDFName + ".pdf";
		}
		
		private void Cleanup()
		{
			TestSuite a = (TestSuite) TestSuite.Current;
			var isCompressedReport = a.ReportCompress;

			if(!isCompressedReport)
			{
				try
				{
					if(System.IO.File.Exists(TestReport.ReportEnvironment.ReportZipFilePath))
					{
						System.IO.File.Delete(TestReport.ReportEnvironment.ReportZipFilePath);
					}
				}
				catch (Exception e)
				{
					Report.Warn("Failed to delete *.rxzlog: " + e.Message);
				}
			}
		}
		
		private string CreatePDFName()
		{
			//Report Status is not part of the ReportName at this stage of the test
			var name = TestReport.ReportEnvironment.ReportName;

			//Get status from TestSuite
			var testsuite = (TestSuite) TestSuite.Current;
			
			if( testsuite.ReportFormatString.Contains("%X"))
			{
				name = name += "_" + GetTestSuiteStatus();
			}

			return name;
		}
		
		private void UpdateError()
		{
			var testsuite = (TestSuite) TestSuite.Current;
			
			if(GetTestSuiteStatus().Contains("Failed"))
			{
				Report.Failure("Rethrow Exception within PDF Module (Necessary for correct error value)");
			}
		}
		
		private string GetTestSuiteStatus()
		{
			string status = "";
			
			var rootChildren = ActivityStack.Instance.RootActivity.Children;
			
			if(rootChildren.Count > 1)
			{
				Console.WriteLine("Multiple TestSuiteActivites, status taken from first entry");
			}
			
			var testSuiteAct = rootChildren[0] as TestSuiteActivity;
			
			if(testSuiteAct != null)
			{
				status = testSuiteAct.Status.ToString();
			}
			
			return status;
		}
		
		public void SendMail()
              {
               
               //String below "RepName" identifies the current TestSuite's report name with the extension of pdf
               string RepName = TestReport.ReportEnvironment.ReportName + ".pdf";
               string ServerHostname = "10.200.12.12";
               string ServerPort = "25";
               
               string Subject = "Test Report: " + RepName;
               string Message = "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=us-ascii\"><font color=\"#ff0000\" size=\"5\">This is an automated message, please do not reply.</font><p>Attached please find the Ranorex Automated Regression Test report.</p><p>If you have any questions, please email: anhua.zhou@nationwideappraisals.com </p><p>Thanks,</p>";
               
               
                try
                {
                	var message = new MailMessage();
                	var fromAddress = new MailAddress("testadmin@nas.com", "Do-Not-Reply"); 
                	message.From = fromAddress;
              		message.To.Add(new MailAddress("testclient@nas.com", "Client Test"));
              		message.To.Add(new MailAddress("testsupport@nas.com", "Support Test"));
              		message.To.Add(new MailAddress("testappraiser@nas.com", "Appraiser Test"));
              		message.Subject = Subject;
              		message.Body = Message;
                	message.IsBodyHtml = true;    
              		
                	//MailMessage mail = new MailMessage (From, To, Subject, Message);
                    SmtpClient smtp = new SmtpClient(ServerHostname, int.Parse(ServerPort));       
                   	smtp.Credentials = CredentialCache.DefaultNetworkCredentials;
                   	smtp.EnableSsl = false;
                    	
                    if(System.IO.File.Exists(RepName))
                    {
                       System.Net.Mail.Attachment attachment;
                       attachment = new System.Net.Mail.Attachment(RepName);
                       message.Attachments.Add(attachment);
                       smtp.Send(message);
                       Report.Success("Email has been sent to '" + message.To.ToString() + "'.");
                    }
                    else
                    {
                       Report.Info("Could not find file: " + RepName);
                    }
                }
                catch(Exception ex)
                {
                    Report.Failure("Mail Error: " + ex.ToString());
                }
            }




	}
}
