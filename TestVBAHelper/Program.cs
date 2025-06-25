using Epr.AradWaterStructures.Request;
using Helper;
using System;
using System.IO;
using System.Runtime.Remoting;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using static Helper.SmsManager.SmsEnum;
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Net;
//using System.Text;
//using System.Threading.Tasks;

namespace TestVBAHelper
{

    class Program
    {
        //static void Main(string[] args)
        //{
        //    ClsMain cls = new ClsMain();
        //	ClsReturn ret;
        //	StringBuilder returnStr = new StringBuilder();
        //	try
        //	{
        //		if (args.Length == 1)
        //		{
        //			ret = cls.BuildAvFile(args[0]);

        //			returnStr.Append($"ret.StatusCode={ret.StatusCode}");
        //			returnStr.Append($"ret.StatusText={ret.StatusText}");

        //			if (ret.StatusCode == 0)
        //			{
        //				returnStr.Append($"ret.Error.Number={ret.Error.Number}");
        //				returnStr.Append($"ret.Error.InnerDescription={ret.Error.InnerDescription}");
        //				returnStr.Append($"ret.Error.OuterDescription={ret.Error.OuterDescription}");
        //			}

        //			File.WriteAllText($"AradAvFileGenerator_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.txt", returnStr.ToString());
        //		}
        //		else
        //			throw new Exception("missing ODBC parameter");

        //	}
        //	catch (Exception ex)
        //	{
        //		File.WriteAllText($"AradAvFileGenerator_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.txt", returnStr.ToString());
        //		MessageBox.Show(ex.Message);
        //	}
        //}

        static void Main()
        {
            bool noEmail = false;
            bool midgam = false;
            bool midgamAuto = false;
            ClsMain cls = new ClsMain();
            ClsReturn ret = cls.SendShovarHovLinkSms(9999,"brngviadev",true);


            //ClsReturn ret = cls.SendShvaHkUnacceptableSms(1010, 309, "Test_050924-123456","בדיקה2", "BrnGviaDev");
            // ClsReturn ret = cls.SendShvaHkUnacceptableSms(1010, 309, "admin_121124-112415", "בדיקה2", "KftGvia_hadracha");
            //ClsReturn ret = cls.GetOnlineRead("10220003", null, "BrnGviaDev");
            //ClsReturn ret2 = cls.GetReadsFile(new DateTime(2023,8,24), null, "BrnGviaDev");
            //ClsReturn ret = cls.RemoveMoneWtr("12121212", "10.08",DateTime.Parse("14/06/23"),1, "BrnGviaDev");
            //ClsReturn ret = cls.CrossMoneWtr("34567", "123456", "BrnGviaDev");

            //ClsReturn ret = cls.ShowMtfProcess(270, "doron", "brngviadev", 277, 277, "0", 0, 0, noEmail, midgam, midgamAuto,10);
            // ClsReturn ret = cls.MtfEmailProcess(270, "doron", "brngviadev", 277, 277, "4059600", 7837, 0,0);
            //	ClsReturn ret = cls.BuildAvFile("");
            //	//ClsReturn ret = cls.GetReadsFile("2023-05-01", null, "BrnGviaDev");
            //	//ClsReturn ret = cls.CrossMoneWtr("123", "456", "brnGviaDev");
            //	//ClsReturn ret = cls.TabuCheckHs(1034, "https://localhost:44406/", "270", "Epr123Epr", "150222134636", "BrnGviaDev");
            //ClsReturn ret = cls.SendInvoiceSms(22222, 2, "05285854", @"d:\hila\33.pdf",true, "KrgGvia_271021_HILA");
            ////	//ClsReturn ret = cls.SendInvoiceSms(22222, 55, "0528585475", @"c:\1\a.pdf", true, "brnGviaDev");
            //	ClsReturn ret = cls.SendGeneralSms("בדיקה יוסף כמות של תווים 1111", "0509671951", "brnGviaDev");
            //	//ClsReturn ret = cls.SendShovarLinkSms(1,2000347, 32394389, "0528585475", "brnGviaDev");
            //	//ClsReturn ret = cls.SendShovarHovLinkSms(1, 6000593, 9999999, "0509671951", "brnGviaDev");
            //	//ClsReturn ret = cls.SendImutPhonePassSms("123", "0528585475", "mvcgvia");
            //	//ClsReturn ret = cls.SendUpdateInfoPassSms("123", "0528585475", "brnGviaDev");
            //	//         Console.OutputEncoding = Encoding.GetEncoding("Windows-1255");
            //	//Console.WriteLine($"Status = {ret.StatusCode}");
            //	//if (ret.StatusCode == 1 || ret.StatusCode < 0)
            //	//	Console.WriteLine($"OuterError = {ret.StatusText}");
            //	//else
            //	//	Console.WriteLine($"InnerError = {ret.Error.InnerDescription}");

            //	//Helper.InvoiceSms.InvoiceSms abc = new Helper.InvoiceSms.InvoiceSms("BrnGviaDev");
            //	//abc.SendInvoiceLinkSms(2, 2, "052858547", @"c:\1\2\33.pdf",ref smsStatus, ref smsStatusDesc, ref errDesc);
            //	//

            //	//ClsMain cls = new ClsMain();
            //	//ClsReturn ret = cls.CreditGuardAPIInquirePaymentTransactionsIntoSql_UsingODBC("BrnGviaDev", "", "", new string[] { "0882814011" }, new DateTime(2012, 1, 1), new DateTime(2022,2, 9), "admin", "7.10.638.6");
            //	//cls.InitializeCreditGuardBrowser(@"C:\Development\TFS\gloabl projects\VBAHelper\Helper\bin\x86\Debug\CGWebPageListener.exe");
            //	////ClsReturn res = cls.CreditGuardWebPageRegisterHK_UsingODBC("BrnGviaDev", "", "", "0882814011", "04/25");  //cls.CreditGuardWebPageVerify_UsingODBC("BrnGviaDev","","", "0882814011","04/25", 034736603);
            //	//ClsReturn res = cls.CreditGuardWebPagePay_UsingODBC("BrnGviaDev", "", "", "test-orik-3", "0882814011", Helper.CreditGuardClient.EnmCreditType.Payments, 1000000.45, 3, false, 666666.67, 0.2);
            //	//cls.ShutDownCreditGuardBrowser();


            //	//ClsMain cls = new ClsMain();
            //	//ClsReturn res = cls.CreditGuardAPIVerify_UsingODBC("BrnGviaDev", "", "", "0882814011", "4557430400000236", "04/24");
            //	//ClsReturn res = cls.CreditGuardAPIRegisterHK_UsingODBC("BrnGviaDev", "", "", "0882814011", "4557430400000236", "04/24");


            //ClsMessage msg = cls.GetEmailMessage(1, "donotreply@eprsystems.co.il", new string[] { "hila@eprsys.com" }, "test subject", "test body");
            //Helper.EmailSender.Settings.ClsMessage ms = cls.getMsg()
            //ClsReturn res = cls.SendSingleEmail("hila@eprsys.com", "test subject", "test body", false, null, 3, "hila", "BrnGviaDev");

            //	//ClsReturn ret = cls.ExecuteHkRegistrationAndHkPayment(1010, 273, "admin","shvasend", "admin_040722", "brnGviaDev");

            //	//ClsReturn res = cls.CreditGuardAPIRegisterHK_FullParams("https://cguat2.creditguard.co.il/xpo/Relay", "epr", "u4?yQjA9", "0882814011", "4444333322221111", "04/25");
            //	//cls.InitializeCreditGuardBrowser(@"E:\Epr\Moaza\MyDev\Current\Helpers\VBAHelper\CGWebPageListener.exe"); //, 0,, "", "", "", "");
            //	//ClsReturn res = cls.CreditGuardWebPageVerify_FullParams(12408, "https://cguat2.creditguard.co.il/xpo/Relay", "epr", "u4?yQjA9", "0882814011", "04/25", 034736603);
            //	//cls.ShutDownCreditGuardBrowser();
            //  ClsReturn res = cls.SendImutPhonePassSms("248554", "0542391219", "BnyGvia");
            //	//ClsReturn ret = cls.ReplaceCreditCards("משתמש3", 2, @"c:\Development\Tasks\CreditCardReplacement\M5867027_0635842_horkeva20211202001403.txt", "brnGviaDev");

            //	//Console.WriteLine($"Status = {ret.StatusCode}");
            //	//if (ret.StatusCode > 0)
            //	//	Console.WriteLine($"OuterError = {ret.StatusText}");
            //	//else
            //	//	Console.WriteLine($"InnerError = {ret.Error.InnerDescription}");



            //	//cls.InitializeCreditGuardBrowser(@"c:\Development\TFS\gloabl projects\VBAHelper\Helper\bin\x86\Debug\CGWebPageListener.exe", 0, 0,"","","","");
            //	//ClsReturn res = cls.CreditGuardWebPageTokenize_UsingOdbc("BrnGviaDev");
            //	//cls.ShutDownCreditGuardBrowser();
            //}


        }
	}
}
