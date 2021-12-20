/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 04/04/2016
 * Time: 12:07 PM
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

namespace RBC_Test
{
    /// <summary>
    /// Description of Login.
    /// </summary>
    [TestModule("4D484909-640D-4E65-BB0D-5BE4276C582C", ModuleType.UserCode, 1)]
    public class Login : ITestModule
    {
        public static RBC_TestRepository repo = RBC_TestRepository.Instance;
		
		static Login instance = new Login();
    	
    	
    			
    	/// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Login()
        {
            // Do not delete - a parameterless constructor is required!
                      
        }

        /// <summary>
        /// Performs the playback of actions in this module.
        /// </summary>
        /// <remarks>You should not call this method directly, instead pass the module
        /// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
        /// that will in turn invoke this method.</remarks>
        /// 
       
        
		public void login(string uid, string pwd, string url)
		{
				Host.Local.OpenBrowser(url, "IE", "", false, false, false, false, false);
				Delay.Milliseconds(400);
				
				repo.RBC_Portal.Front.welcome.Click();
				Delay.Milliseconds(100);
			
				repo.RBC_Portal.Front.username.PressKeys(uid);
				repo.RBC_Portal.Front.password.PressKeys(pwd);
				repo.RBC_Portal.Front.submit.Click();
				Delay.Milliseconds(200);
				
				/*
				if (repo.RBC_Nas.NotificationInfo.Exists())
				{
					repo.RBC_Nas.NotForThisSite.Click();
				}
				*/
			
				Report.Log(ReportLevel.Success, "Validation", "RBC user: [" + uid + "]" + " log in successfully.");
				Validate.Exists(repo.RBC_Portal.QuickNavbarTopRight.QuickIcon);
							
		}
		void ITestModule.Run()
        {
            Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 100;
            Delay.SpeedFactor = 1.0;
           
        }
    }
}
