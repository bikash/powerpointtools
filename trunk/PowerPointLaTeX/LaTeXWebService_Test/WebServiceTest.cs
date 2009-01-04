using LaTeXWebService;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace LaTeXWebService_Test
{
    
    
    /// <summary>
    ///This is a test class for WebServiceTest and is intended
    ///to contain all WebServiceTest Unit Tests
    ///</summary>
    [TestClass()]
    public class WebServiceTest
    {


        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        /// <summary>
        ///A test for getURLData
        ///</summary>
        [TestMethod()]
        [DeploymentItem("LaTeXWebService.dll")]
        public void getURLDataTest()
        {
            WebService_Accessor target = new WebService_Accessor(); // TODO: Initialize to an appropriate value
            string url = "http://home.in.tum.de/~kirschan/"; // TODO: Initialize to an appropriate value
            WebService.URLData actual;
            actual = target.getURLData(url);
            System.Windows.Forms.MessageBox.Show(System.Text.Encoding.ASCII.GetString(actual.content));
        }

        /// <summary>
        ///A test for getRequestURL
        ///</summary>
        [TestMethod()]
        [DeploymentItem("LaTeXWebService.dll")]
        public void getRequestURLTest()
        {
            WebService_Accessor target = new WebService_Accessor(); // TODO: Initialize to an appropriate value
            string latexCode = string.Empty; // TODO: Initialize to an appropriate value
            string expected = string.Empty; // TODO: Initialize to an appropriate value
            string actual;
            actual = target.getRequestURL(latexCode);
            Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for compileLaTeX
        ///</summary>
        [TestMethod()]
        public void compileLaTeXTest()
        {
            WebService target = new WebService(); // TODO: Initialize to an appropriate value
            string latexCode = @"a \le b"; // TODO: Initialize to an appropriate value
            WebService.URLData expected = new WebService.URLData(); // TODO: Initialize to an appropriate value
            WebService.URLData actual;

            actual = target.compileLaTeX(latexCode);
            MemoryStream stream = new MemoryStream(actual.content, false);
 
            //TestImage window = new TestImage();
            //window.picture.Image = Image.FromStream(stream);
            //window.ShowDialog();
            TestImage window = new TestImage();
            window.picture.Image = Image.FromStream(stream);
            window.ShowDialog();
        }

        /// <summary>
        ///A test for WebService Constructor
        ///</summary>
        [TestMethod()]
        public void WebServiceConstructorTest()
        {
            WebService target = new WebService();
            Assert.Inconclusive("TODO: Implement code to verify target");
        }
    }
}
