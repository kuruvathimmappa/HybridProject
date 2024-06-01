package orangehrm.testcases;
import org.testng.annotations.Test;
import orangehrm.library.Employee;
import orangehrm.library.LoginPage;
import utils.AppUtils;
import utils.XLUtils;

public class OrangeHRM_HybridTest extends AppUtils
{
	String xlfile="C:\\Selenium\\orangeHRM_Hybrid_DDT_Test\\datafiles\\OrangeHRM_Hybrid_DDT_testt.xlsx";
	String tcsheet="TestCases";
	String tssheet="TestSteps";
	
	@Test
	public void CheckOrangeHRM() throws Exception
	{
		
		int tccount=XLUtils.getRowCount(xlfile, tcsheet);
		int tscount=XLUtils.getRowCount(xlfile, tssheet);
		String tcid,tcexeflag,tstcid,keyword;
		LoginPage lp=new LoginPage();
		Employee emp=new Employee();
		String uid,pwd;
		String fname,mname,lname;
		boolean res = false;
		for(int i=1;i<=tccount;i++) 
		{
			tcid=XLUtils.getStringData(xlfile, tcsheet, i, 0);
			tcexeflag=XLUtils.getStringData(xlfile, tcsheet, i, 2);
							
			if(tcexeflag.equalsIgnoreCase("y"))
			{
				for(int j=1;j<=tscount;j++)
				{
					tstcid=XLUtils.getStringData(xlfile, tssheet, j, 0);
					if(tstcid.equalsIgnoreCase(tcid))
					{
						keyword=XLUtils.getStringData(xlfile, tssheet, j, 4);
						switch(keyword.toLowerCase())
						{
 						case "adminlogin":
							uid=XLUtils.getStringData(xlfile, tssheet, j, 5);
							pwd=XLUtils.getStringData(xlfile, tssheet, j, 6);
							lp.login(uid, pwd);
							 res=lp.isAdminModuleDisplayed();
//							Assert.assertTrue(res);
							break;
						case "logout":
							res=lp.logout();
							break;
						case "invalidlogin":
							uid=XLUtils.getStringData(xlfile, tssheet, j, 5);
							pwd=XLUtils.getStringData(xlfile, tssheet, j, 6);
							lp.login(uid, pwd);
							res=lp.isErrMsgDisplayed();
							break;
						case "empreg":
							fname=XLUtils.getStringData(xlfile, tssheet, j, 5);
							lname=XLUtils.getStringData(xlfile, tssheet, j, 6);
							emp.addEmployee(fname, lname);
							break;
						}
						String step_result;
						if(res)
						{
							step_result="Pass";
							XLUtils.setData(xlfile, tssheet, j, 3, step_result);
							XLUtils.fillGreenColor(xlfile, tssheet, j, 3);
						}else
						{
							step_result="Fail";
							XLUtils.setData(xlfile, tssheet, j, 3, step_result);
							XLUtils.fillRedColor(xlfile, tssheet, j, 3);
						}
						String stepcaseresult;
						stepcaseresult= XLUtils.getStringData(xlfile, tcsheet, i, 3);
						if(!stepcaseresult.equalsIgnoreCase("Fail"))
						{	
						XLUtils.setData(xlfile, tcsheet, i, 3, step_result);
						}
						stepcaseresult= XLUtils.getStringData(xlfile, tcsheet, i, 3);
						if(stepcaseresult.equalsIgnoreCase("Pass"))
						{
							XLUtils.fillGreenColor(xlfile, tcsheet, i, 3);
						}else
						if(stepcaseresult.equalsIgnoreCase("Fail"))
						{
							XLUtils.fillRedColor(xlfile, tcsheet, i, 3);
						}
					}
				}
			}
			else
			{
				XLUtils.setData(xlfile, tcsheet, i, 3, "Blocked");
				XLUtils.fillRedColor(xlfile, tcsheet, i, 3);
			}
		}
	}
}
