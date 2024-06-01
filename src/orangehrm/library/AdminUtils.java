package orangehrm.library;

import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;

import utils.AppUtils;

public class AdminUtils extends AppUtils
{
	LoginPage lp;
	String adminuid = "Admin";
	String adminpwd = "Qedge123!@#";
	
	@BeforeTest
	public void adminLogin()
	{
		lp = new LoginPage();
		lp.login(adminuid, adminpwd);
	}
	
	@AfterTest
	public void adminLogout()
	{
		lp.logout();
	}
	
}
