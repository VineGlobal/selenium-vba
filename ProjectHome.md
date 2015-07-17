Selenium VBA is a Windows COM library that uses the popular [Selenium](http://docs.seleniumhq.org/) web testing tool.<br>
It makes it possible to automate web browsing using Excel and VBA code or using a simple VBS file executed by double-clicking on it.<br>
User's actions can be recorded using the Firefox plugin named "Selenium IDE" and translated to VBA or VBS with the provided formatters.<br>

<blockquote><img src='http://selenium-vba.googlecode.com/svn/wiki/schematic.png' /></blockquote>

<b>It comes usefull to :</b>
<ul><li>Automate repetitive web browser tasks.<br>
</li><li>Quickly fill a web form multiple times with an Excel data set.<br>
</li><li>Extract data from a web page in an Excel sheet.<br>
</li><li>Run web tests against an Excel data set (Data-Driven Testing).<br>
</li><li>Take screenshots of a Web site and save them in a PDF file.<br>
</li><li>Use the selenium automation framework within QTP (Quick Test Pro).<br>
</li><li>Compare the rendering of two web pages to quickly detect regressions<br>
</li><li>Measure the page loading time as well as the server response time.</li></ul>

<b>This project provides :</b>
<ul><li>A COM library to use Selenium with Excel in the Visual Basic Editor or within a visual basic script (VBS).<br>
</li><li>The Selenium IDE plugin with new formatters to convert recorded actions to the VBA/VBS programming language.<br>
</li><li>All the Selenium 2 and Selenium 1(RC) commands.<br>
</li><li>A console application to run scripts<br>
</li><li>A simple and quick way to drive Firefox, IE, Chrome, Safari and PhantomJS (headless webkit).</li></ul>

<h3>Download</h3>
<ul><li><a href='https://bintray.com/florentbr/generic/selenium-vba/view/files?order=desc'>Latest version</a>
</li><li><a href='https://selenium-vba.googlecode.com/svn/trunk/wrapper/ChangeLog.txt'>Change Log</a></li></ul>

<h3>Documentation</h3>
There are multiple ways to get documentation:<br>
<ul><li>From the <a href='http://docs.seleniumhq.org/docs/'>Selenium website</a> with samples referring to the CShap language (Most of the commands have same syntax)<br>
</li><li>From the provided help file : Programs>SeleniumWrapper>API Documentation<br>
</li><li>Using the object browser in the VBA editor (F2)<br>
</li><li>From the web with the keyword "Selenium CShap" followed by your issue</li></ul>

<h3>Minimum Requirements</h3>
<ul><li>Windows XP<br>
</li><li>Microsoft Office 2003/2007/2010/2013 with option "VBA For Application" installed <a href='http://www.microsoft.com/en-us/download/details.aspx?id=10624'>[Update required for Office 2003</a>]<br>
</li><li>Mozilla Firefox 8 to 34 <a href='https://ftp.mozilla.org/pub/mozilla.org/firefox/releases'>[Download Page</a>]<br>
</li><li>Microsoft .Net Framework 3.5 SP1 <a href='http://www.microsoft.com/en-us/download/confirmation.aspx?id=22'>[Download Page</a>]</li></ul>

<h3>How to create/record a VBA script ?</h3>
<ul><li>Launch Firefox and Selenium IDE (Ctrl+Alt+S)<br>
</li><li>Open or record a script<br>
</li><li>Click on Menu "Option>Format>VBA", select the Tab "Source" and copy the text<br>
</li><li>Or Click on Menu "File>Export Test Case As>VBA" and save the file<br>
</li><li>Paste the generated script in a module in Excel VBA (Alt + F11)</li></ul>

<blockquote><img src='http://selenium-vba.googlecode.com/svn/wiki/screen-a.png' /></blockquote>

<h3>How to run a VBA script in Excel with Visual Basic for Application ?</h3>
<ul><li>Open a new workbook or the provided template (All Programs/SeleniumWrapper/Excel Template)<br>
</li><li>Click on Menu "Tools>Macro>Visual Basic Editor"<br>
</li><li>Add the reference "SeleniumWraper Type Library" in Tools>References (Already present in the template)<br>
</li><li>In VBE, click on menu "Insert>Module" and paste your code<br>
</li><li>Click on Run in VBE or Run the macro from Excel</li></ul>

<blockquote><img src='http://selenium-vba.googlecode.com/svn/wiki/screen-b.png' /></blockquote>

<h3>How to select the browser to work with?</h3>
This application works with real browsers, which makes it possible to interact with dynamic contents created by Javascript.<br>
The command to lauch a browser requires le name of the browser and the root of the web site.<br>
<blockquote>Here is an example which launches firefox, opens the science page of Yahoo news, the politics page and closes the browser :<br>
<pre><code>Public Sub TC_Browsers()<br>
   Dim driver As New SeleniumWrapper.WebDriver<br>
   driver.start "firefox", "http://news.yahoo.com"<br>
   driver.open "/science"<br>
   driver.open "/politics"<br>
   driver.stop<br>
End Sub<br>
</code></pre>
There are 4 browsers that can be controlled locally that need to be installed to work with:<br>
</blockquote><ul><li>Firefox : "firefox"<br>
</li><li>Chrome : "chrome"<br>
</li><li>InternetExplorer : "ie"<br>
</li><li>InternetExplorer 64 bits : "ie64"<br>
</li><li>Safari : "safari"<br>
There is also one headless browser(PhantomJS) included in the installation which behave like a browser but without a window to interect with:<br>
</li><li>PhantomJS : "phantomjs"</li></ul>

Another feature is the possibility to control a browser remotely using the the remote web driver server: <a href='https://code.google.com/p/selenium/wiki/RemoteWebDriverServer'>https://code.google.com/p/selenium/wiki/RemoteWebDriverServer</a>

<blockquote>Here is a example to remotely control Firefox on the station 192.168.0.26 :<br>
<pre><code>Public Sub TC_Browsers()<br>
   Dim driver As New SeleniumWrapper.WebDriver<br>
   driver.startRemotely "firefox", "http://192.168.0.26:4444/wd/hub" "http://news.yahoo.com"<br>
   driver.open "/science"<br>
   ...<br>
End Sub<br>
</code></pre></blockquote>

<h3>How to get an element locator ?</h3>
<h5>Solution 1 :</h5>
Using Selenium IDE, copy the recorded selector in the field "Target".<br>
This selector can then be used as follow :<br>
selenium.click "target"<br>
selenium.type "target", "your text"<br>
For the a Selenium 2 command, remove the prefix of the target and use the corresponding command:<br>
<ul><li>Target "css=value" : driver.findElementByCssSelector("value").sendKeys "your text"<br>
</li><li>Target "xpath=value" : driver.findElementByXPath("value").sendKeys "your text"<br>
</li><li>Target "name=value" : driver.findElementByName("value").sendKeys "your text"<br>
</li><li>Target "id=value" : driver.findElementById("value").sendKeys "your text"<br>
</li><li>Target "link=value" : driver.findElementByLinkText("value").sendKeys "your text"</li></ul>

<h5>Solution 2 :</h5>
In Firefox, right click on the element and click on inspect element<br>
Right click one the node and click "Copy Unique Selector"<br>
Then to use the selector:<br>
selenium.click "css=selector"<br>
driver.findElementByCssSelector("selector").click<br>
<br>
<h3>Excel script example: Web automation</h3>
<blockquote>This example runs a simple web search and pastes a screenshot of the result in a worksheet:<br>
<pre><code>Public Sub TC001()<br>
   Dim selenium As New SeleniumWrapper.WebDriver<br>
   selenium.start "firefox", "http://www.google.com/"<br>
   selenium.open "/"<br>
   selenium.type "name=q", "Eiffel tower"<br>
   selenium.click "name=btnG"<br>
   selenium.wait 1000<br>
   selenium.getScreenshot().copy<br>
   selenium.stop<br>
   Sheets(1).Range("A10").PasteSpecial       'Paste the screenshoot at range A10<br>
End Sub<br>
</code></pre></blockquote>

<h3>Excel script example: Data driven testing</h3>
<blockquote>In this example, a worksheet contains a list of Urls and expected titles.<br>
Each Url is oppend in Firefox, then the web page title is compared with the expected one and the result is sent back to the worksheet.<br>
<pre><code>Public Sub TC002()<br>
    Dim selenium As New SeleniumWrapper.WebDriver, r As Range<br>
    selenium.Start "firefox", "http://www.google.com"       'Launch Firefox<br>
    For Each r In Range("MyValues").Rows                    'Loop for each row in the range named "MyValues"<br>
        selenium.open r.Cells(1, 1).Text                    'open the link defined in the first column of "MyValues"<br>
        selenium.waitForNotTitle ""                         'wait for the title to load<br>
        r.Cells(1, 3) = selenium.verifyTitle(r.Cells(1, 2))   'Compare the page title with the second column and write the verification result in the third column<br>
    Next<br>
    selenium.stop<br>
End Sub<br>
</code></pre>
This wrapper implements assertion and verification commands.<br>
When an assert fails (ex: assertTitle), an error popup appears and the execution is stopped.<br>
A verification command (ex: verifyTitle) just returns a string with the test result without breaking.</blockquote>

<h3>VBS Script example : Create a PDF of screenshots</h3>
<blockquote>This scrip will help you to quickly inspect a list of web pages without having to navigate the website and wait for page loading.<br>
It automatically opens the webpage, takes a screenshot and saves it in a Pdf file:<br>
<pre><code>Set selenium = CreateObject("SeleniumWrapper.WebDriver")<br>
Set pdf = CreateObject("SeleniumWrapper.PdfFile")<br>
selenium.start "firefox", "http://www.google.com"<br>
selenium.open "search?q=eiffel+tower"<br>
pdf.addImage selenium.getScreenshot(), "Google search - Eiffel tower"<br>
selenium.open "http://maps.google.com/maps?q=eiffel+tower"<br>
pdf.addImage selenium.getScreenshot(), "Google map - Eiffel tower"<br>
pdf.saveAs "c:\selenium-capture.pdf"<br>
selenium.stop<br>
</code></pre></blockquote>

<h3>Excel script example: Web scraping (From version 1.0.18)</h3>
<blockquote>This example gets all the world market indexes in the first worksheet and all the top stories tiltles in the second one:<br>
<pre><code>Public Sub TC003()<br>
    Dim driver As New SeleniumWrapper.WebDriver<br>
    driver.Start "chrome", "https://www.google.co.uk"  'Starts the browser<br>
    driver.Open "/finance"  'Opens the finance page<br>
    <br>
    Dim data1, data2<br>
    data1 = driver.findElementByCssSelector("#markets table").AsTable.GetData() 'Gets the world market indexes from the table<br>
    Sheet1.[A1].Resize(UBound(data2, 1), UBound(data2, 2)).Value = data1  'Writes the collected data in the first worksheet<br>
    data2 = driver.findElementsByCssSelector("#market-news-stream a.title").GetData()   'Gets the top stories titles<br>
    Sheet2.[A1].Resize(UBound(data2)).Value = data2 'Writes the collected data in the second worksheet<br>
    <br>
    driver.stop 'Stops the browser<br>
End Sub<br>
</code></pre></blockquote>

<h3>VBS Script example : Compare the rendering of two web pages (From version 1.0.18)</h3>
<blockquote>Getting tired playing "Where is Wally?" for each delivery ?<br>
Here is an example to quickly identify changes between two versions of a web page :<br>
<pre><code>Set driver = CreateObject("SeleniumWrapper.WebDriver")<br>
driver.start "firefox", "http://www.google.co.uk" 'Starts the browser<br>
driver.open "/" 'Opens google search version UK<br>
Set imageA = driver.getScreenshot() 'Captures the rendering<br>
driver.open "http://www.google.fr" 'Opens Google search version FR<br>
Set imageB = driver.getScreenshot()  'Captures the rendering<br>
driver.stop 'Stops the browser<br>
imageA.compareTo(imageB).saveAs "diff.png" 'Compares images and saves the result<br>
</code></pre>
For result, an image showing differences with a non-black colour : <br>
<img src='http://selenium-vba.googlecode.com/svn/wiki/image-comp.png' /></blockquote>


<h3>Excel script example : Measure real loading times (From version 1.0.18)</h3>
The getPerformanceTiming command returns latency-related performance informations :<br>
<ul><li>Page load : Total page load delay experienced by the user<br>
</li><li>Redirect : Time taken for page redirection<br>
</li><li>DNS :  time taken to perform DNS lookup to the server<br>
</li><li>Connecting :  time taken to connect to the server<br>
</li><li>Waiting : time taken for the server to start responding<br>
</li><li>Receiving : time taken to receive the data from the server<br>
</li><li>DOM : Time spent building the DOM<br>
</li><li>Events : Time taken to handle onLoad event</li></ul>

<blockquote>This example navigates through Yahoo news and writes the metrics for each web page in an Excel WorkSheet:<br>
<pre><code>Function NextRow() As Range  <br>
    Static i As Integer<br>
    i = IIf(i = 0, 1, i + 1)	'Increments the row index<br>
    Set NextRow = Range("A:I").Rows(i)	'Returns the next row<br>
End Function<br>
<br>
Public Sub GetPerformanceTiming()<br>
    Dim driver As New SeleniumWrapper.WebDriver<br>
    driver.Start "firefox", "http://uk.news.yahoo.com"	'Starts firefox<br>
    <br>
    'Writes the titles on the first row<br>
    NextRow() = Array("Page Url", "Page load", "Redirect", "DNS", "Connecting", "Waiting", "Receiving", "DOM", "Events")<br>
    <br>
    'Opens each page and copies the metrics in the WorkSheet ( 1 row for each page opened)<br>
    driver.Open "tech"<br>
      NextRow() = driver.getPerformanceTiming()<br>
    driver.Open "world"<br>
      NextRow() = driver.getPerformanceTiming()<br>
    driver.Open "opinion"<br>
      NextRow() = driver.getPerformanceTiming()<br>
    driver.Open "business"<br>
      NextRow() = driver.getPerformanceTiming()<br>
    <br>
    driver.stop	'Closes Firefox<br>
End Sub<br>
</code></pre></blockquote>


<h3>Third-party software components included in the installation package</h3>
<ul><li>.NET Selenium Client Drivers <a href='http://seleniumhq.org'>[Home page</a>]<br>
</li><li>IE driver server <a href='http://seleniumhq.org'>[Home page</a>]<br>
</li><li>Chrome driver server <a href='http://code.google.com/p/chromedriver'>[Home page</a>]<br>
</li><li>PhantomJS server <a href='http://phantomjs.org/'>[Home page</a>]<br>
</li><li>PDFsharp library <a href='http://www.pdfsharp.com/PDFsharp/'>[Home page</a>]</li></ul>

<h3>Tested environments</h3>
<ul><li>Win7 / Excel 2010/ Firefox 11 / Selenium IDE 1.7.1<br>
</li><li>WinXP / Excel 2003 / Firefox 10 / Selenium IDE 1.6.0</li></ul>

<h3>Author</h3>
<blockquote>Florent BREHERET