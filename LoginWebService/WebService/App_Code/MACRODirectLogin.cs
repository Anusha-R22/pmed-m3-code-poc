using System;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Text;
using DirectLogin30;

[WebService(Namespace = "http://localhost/MACRO30DirectLogin/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
public class MACRODirectLogin : System.Web.Services.WebService
{
    public MACRODirectLogin () {

        //Uncomment the following line if using designed components 
        //InitializeComponent(); 
    }

    [WebMethod]
    public int Login(string userName, string passWord, out string userXML, out string errorXML)
    {
        int loginResult = -1;
        userXML = "";
        errorXML = "";

        // perform login
        try
        {
            DirectLoginClass loginDirect = new DirectLoginClass();
            loginResult = (int)loginDirect.Login(userName, passWord, ref userXML);
        }
        catch (Exception ex)
        {
            errorXML = FormatErrorXml(ex);
            return 2; // fail code
        }

        return loginResult;
    }

    /*
    [WebMethod]
    public string HelloWorld()
    {
        return "Hello World";
    }
    */

    private static string FormatErrorXml(Exception ex)
    {
        StringBuilder sbError = new StringBuilder();
        sbError.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        sbError.Append("<error id=\"not defined\">");
        sbError.Append("<description>");
        sbError.Append(ex.Message.ToString());
        sbError.Append("</description>");
        sbError.Append("<source>not defined</source>");
        sbError.Append("<data>not defined</data>");
        sbError.Append("</error>");
        return sbError.ToString();
    }
}
