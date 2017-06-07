package windowApplications;

import org.sikuli.script.FindFailed;
import org.sikuli.script.Key;
import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class OpenApplications {
	/* Extending the project with extent report */
	ExtentReports extent;
	ExtentTest extentTest;
	String extentReportFile = System.getProperty("user.dir") + "\\extentReportFile.html";
	/*----------------------------------------*/

	String path = "D:\\amit\\Java_programs\\StartingWindowsApplication\\libs\\";

	Screen screen = new Screen();

	@Test(priority = 0)
	public void openWord() {
		try {
			extentTest = extent.startTest("Opening Microsoft Document!",
					"This report verifies the opening of microsoft word document");
			extentTest.log(LogStatus.INFO, "Start icon is clisked");
			screen.click(patterns("start"));
			extentTest.log(LogStatus.INFO, "Word is typed to search the program.");
			screen.type(patterns("enterProgramName"), "word");
			extentTest.log(LogStatus.INFO, "Word is clicked.");
			screen.click(patterns("word"));
			extentTest.log(LogStatus.INFO, "Click the blank document to open the blank page for word sheet.");
			screen.click(patterns("word_blankdocument"));
		} catch (FindFailed failed) {
			extentTest.log(LogStatus.FAIL, "Failed");
			extent.endTest(extentTest);
		} finally {
			extent.endTest(extentTest);
		}

	}

	@Test(priority = 1)
	public void openNotePad() {
		try {
			extentTest = extent.startTest("Opening Notepad!", "This report verifies the opening of notepad");
			extentTest.log(LogStatus.INFO, "Start icon is clicked");
			screen.click(patterns("start"));
			extentTest.log(LogStatus.INFO, "Notepad is typed to search the program.");
			screen.type(patterns("enterProgramName"), "notepad");
			extentTest.log(LogStatus.INFO, "Notepad is clicked.");
			screen.click(patterns("notepad"));
		} catch (FindFailed failed) {
			extentTest.log(LogStatus.FAIL, "Failed");
			extent.endTest(extentTest);
		} finally {
			extent.endTest(extentTest);
		}

	}

	@Test(priority = 2)
	public void openChrome() {
		try {
			extentTest = extent.startTest("Opening Chrome!", "This report verifies the opening of chrome");
			extentTest.log(LogStatus.INFO, "Start icon is clicked");
			screen.click(patterns("start"));
			extentTest.log(LogStatus.INFO, "chrome is typed to search the program.");
			screen.type(patterns("enterProgramName"), "chrome");
			extentTest.log(LogStatus.INFO, "Chrome is clicked.");
			screen.click(patterns("chrome"));
		} catch (FindFailed failed) {
			extentTest.log(LogStatus.FAIL, "Failed");
			extent.endTest(extentTest);
		} finally {
			extent.endTest(extentTest);
		}
	}

	@Test(priority = 3)
	public void openExcel() {
		try {
			extentTest = extent.startTest("Opening Microsoft Excel!",
					"This report verifies the opening of Microsoft Excel");
			extentTest.log(LogStatus.INFO, "Start icon is clicked");
			screen.click(patterns("sdstart"));
			extentTest.log(LogStatus.INFO, "excel is typed to search the program.");
			screen.type(patterns("enterProgramName"), "escel");
			extentTest.log(LogStatus.INFO, "Excel is clicked.");
			screen.click(patterns("excel"));
			extentTest.log(LogStatus.INFO, "Click the blank document to open the blank page for excel sheet.");
			screen.click(patterns("excel_blankdocument"));
		} catch (FindFailed failed) {
			extentTest.log(LogStatus.FAIL, failed.getLocalizedMessage());
			extentTest.log(LogStatus.FAIL, failed.getCause());
			extentTest.log(LogStatus.FAIL, failed.getMessage());
			extent.endTest(extentTest);
		} finally {
			extent.endTest(extentTest);
		}

	}

	@Test(priority = 4)
	public void openWordPad() {
		try {
			extentTest = extent.startTest("Opening Wordpad!", "This report verifies the opening of Wordpad");
			extentTest.log(LogStatus.INFO, "Start icon is clicked");
			screen.click(patterns("start"));
			extentTest.log(LogStatus.INFO, "wordpad is typed to search the program.");
			screen.type(patterns("enterProgramName"), "wordpad");
			extentTest.log(LogStatus.INFO, "Wordpad is clicked.");
			screen.click(patterns("wordpad"));
		} catch (FindFailed failed) {
			extentTest.log(LogStatus.FAIL, "Failed");
			extent.endTest(extentTest);
		} finally {
			extent.endTest(extentTest);
		}
	}

	@Test(priority = 5)
	public void openPOSTMAN() {
		try {
			extentTest = extent.startTest("Opening POSTMAN!", "This report verifies the opening of POSTMAN");
			extentTest.log(LogStatus.INFO, "Start icon is clicked");
			screen.click(patterns("start"));
			extentTest.log(LogStatus.INFO, "postman is typed to search the program.");
			screen.type(patterns("enterProgramName"), "postman");
			extentTest.log(LogStatus.INFO, "POSTMAN is clicked.");
			screen.click(patterns("postman"));
		} catch (FindFailed failed) {
			extentTest.log(LogStatus.FAIL, "Failed");
			extent.endTest(extentTest);
		} finally {
			extent.endTest(extentTest);
		}
	}

	@Test(priority = 6)
	public void openWAMP() {
		try {
			extentTest = extent.startTest("Opening WAMP!", "This report verifies the opening of WAMP");
			extentTest.log(LogStatus.INFO, "Start icon is clicked");
			screen.click(patterns("start"));
			extentTest.log(LogStatus.INFO, "wamp is typed to search the program.");
			screen.type(patterns("enterProgramName"), "wamp");
			extentTest.log(LogStatus.INFO, "Wampserver64 is clicked.");
			screen.click(patterns("wamp"));
		} catch (FindFailed failed) {
			extentTest.log(LogStatus.FAIL, "Failed");
			extent.endTest(extentTest);
		} finally {
			extent.endTest(extentTest);
		}

	}

	@Test(priority = 7)
	public void openOutlook() {
		try {
			extentTest = extent.startTest("Opening Outlook!", "This report verifies the opening of Outlook");
			extentTest.log(LogStatus.INFO, "Start icon is clicked");
			screen.click(patterns("start"));
			extentTest.log(LogStatus.INFO, "outlook is typed to search the program.");
			screen.type(patterns("enterProgramName"), "outlook");
			extentTest.log(LogStatus.INFO, "Outlook 2013 is clicked.");
			screen.click(patterns("outlook"));
			screen.click(patterns("outlookpassword"));
			extentTest.log(LogStatus.INFO, "Selecting all text in the password alert.");
			screen.type("A", Key.CTRL);
			extentTest.log(LogStatus.INFO, "Deleting the previous password from the password field.");
			screen.type(Key.BACKSPACE);
			extentTest.log(LogStatus.INFO, "Entering the password in the password field.");
			screen.paste("Googl@123");
			screen.click(patterns("outlookOk"));
		} catch (FindFailed failed) {
			extentTest.log(LogStatus.FAIL, "Failed");
			extent.endTest(extentTest);
		} finally {
			extent.endTest(extentTest);
		}
	}

	public Pattern patterns(String string) {
		Pattern pattern = new Pattern(path + string + ".PNG");
		return pattern;
	}

	@BeforeTest
	public void startUp() {
		extent = new ExtentReports(extentReportFile, false);
	}

	@AfterTest
	public void flush() {
		extent.endTest(extentTest);
		extent.flush();
	}
}
