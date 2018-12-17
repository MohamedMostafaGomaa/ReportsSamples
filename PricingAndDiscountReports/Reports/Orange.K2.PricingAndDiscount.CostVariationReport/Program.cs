using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Net.Mail;
using Microsoft.Reporting.WebForms;
using Orange.K2.Common.Utilities;
using System.Configuration;
using System.IO;
using System.Globalization;
using Microsoft.Practices.Unity;




[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace Orange.K2.PricingAndDiscount.CostVariationReport
{
    class Program
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static void Main(string[] args)
        {
            logger.Error("******************************************* Start *******************************************");
            logger.Error("Start Cost Variation Service at : " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));

            try
            {

                var resolver = UnityConfig.GetConfiguredContainer();
                var _Utilities = resolver.Resolve<IUtilities>();

                ReportViewer CostVariationReport = new ReportViewer();
                CostVariationReport.ProcessingMode = ProcessingMode.Remote;
                CostVariationReport.ServerReport.ReportServerUrl = new Uri(_Utilities.GetAppSetting("ReportServerUrl"));
                CostVariationReport.ServerReport.ReportPath = _Utilities.GetAppSetting("CostVariationReportReportPath");
                CostVariationReport.ServerReport.ReportServerCredentials = new ReportServerCredentials(_Utilities.GetAppSetting("ReportServerUsername"), _Utilities.GetAppSetting("ReportServerPassword"), _Utilities.GetAppSetting("ReportServerDomain"));
                AssignReportParameters(CostVariationReport);
                string mimeType;
                string encoding;
                string extension;
                string format = "EXCEL";
                string devInfo = null;
                string[] streams;
                byte[] result = null;
                Microsoft.Reporting.WebForms.Warning[] warnings;
                result = CostVariationReport.ServerReport.Render(format, devInfo, out mimeType, out encoding, out extension, out streams, out warnings);

                string path = ConfigurationManager.AppSettings["FilePath"].ToString();

                using (System.IO.FileStream writer = new FileStream(path, FileMode.Create))
                {
                    logger.Error("Report Created at : " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));
                    writer.Write(result, 0, result.Length);
                    logger.Error("Result written to the Report at : " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));
                    writer.Close();
                }

                int hours = int.Parse(System.Configuration.ConfigurationManager.AppSettings["Hour"]);
                int minutes = int.Parse(System.Configuration.ConfigurationManager.AppSettings["minutes"]);

                DateTime CureentDate = DateTime.Now;
                DateTime tempDate = CureentDate.AddMonths(1);
                DateTime NextDate = new DateTime();
                if (CureentDate.Day == 1 && CureentDate.Hour < 7)
                {
                    NextDate = new DateTime(CureentDate.Year, CureentDate.Month, 1, hours, minutes, 0);
                }
                else
                {
                    NextDate = new DateTime(tempDate.Year, tempDate.Month, 1, hours, minutes, 0);
                }
                TimeSpan TimeDiff = NextDate.Subtract(DateTime.Now);
                string schedule = string.Format("{0} day(s) {1} hour(s) {2} minute(s) {3} second(s)", TimeDiff.Days, TimeDiff.Hours, TimeDiff.Minutes, TimeDiff.Seconds);
                logger.Error("Cost Variation  Service scheduled to run after: " + schedule + "From" + " {0} " + " Next Date To Execute Is: " + NextDate.ToString("dd/MM/yyyy hh:mm:ss tt"));

                SendEmail();
                CostVariationReport.ServerReport.Refresh();

                logger.Error("******************************************* End *******************************************");
            }
            catch (Exception exp)
            {
                logger.Error("Exception Error : " + exp.Message);

            }

        }
        private static void AssignReportParameters(ReportViewer CostVariationReport)
        {

            try
            {

                DateTime now = DateTime.Now;
                DateTime realDate = now.AddMonths(-1);
                var startDate = new DateTime(realDate.Year, realDate.Month, 1);
                var endDate = startDate.AddMonths(1).AddDays(-1);
                //var startDate = new DateTime(realDate.Year, 10, 1);
                //var endDate = startDate.AddMonths(1).AddDays(-1);
                List<Microsoft.Reporting.WebForms.ReportParameter> parameterList = new List<Microsoft.Reporting.WebForms.ReportParameter>();
                parameterList.Add(new Microsoft.Reporting.WebForms.ReportParameter("DateFrom", startDate.ToString()));
                parameterList.Add(new Microsoft.Reporting.WebForms.ReportParameter("DateTo", endDate.ToString()));
                CostVariationReport.ServerReport.SetParameters(parameterList);
                logger.Error("Report For Month : " + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(realDate.Month) + "  " + realDate.Year);
                logger.Error("Report Start Date : " + startDate.ToString("dd/MM/yyyy hh:mm:ss tt"));
                logger.Error("Report End Date : " + endDate.ToString("dd/MM/yyyy hh:mm:ss tt"));
            }
            catch (Exception exp)
            {
                logger.Error("Exception Error in Adding Report Parameters : " + exp.Message);

            }

        }

        private static void SendEmail()
        {
            try
            {

                var resolver = UnityConfig.GetConfiguredContainer();
                var _Utilities = resolver.Resolve<IUtilities>();
                string path = _Utilities.GetAppSetting("FilePath");
                string Sender = _Utilities.GetAppSetting("Sender");
                string SenderUsername = _Utilities.GetAppSetting("SenderUsername");
                string SenderPassword = _Utilities.GetAppSetting("SenderPassword");
                string AllReceiver = _Utilities.GetAppSetting("Receiver");
                string AllReceiverName = _Utilities.GetAppSetting("ReceiverName");
                string[] Receiver = AllReceiver.Split(',');
                string[] ReceiverName = AllReceiverName.Split(',');

                string mailBody = _Utilities.GetAppSetting("Mailbody");
                XmlDocument document = new XmlDocument();
                document.Load(mailBody);
                string body = document.InnerText.ToString();//.GetElementsByTagName("mbody")
                                                            //var body = document.Root.Elements().Select(x => x.Element("mbody").ToString());
                                                            //string.Format("Dears: <br/><br/>kindly You find monthly cost variation report for : <b> {0} {1} </b> attached <br/><br/><b>Note:</b>This is an automated mail. Do not reply to this mail<br/><br/>Regards <br/> Orange System", CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(realDate.Month), realDate.Year);
                for (int i = 0; i < Receiver.Length; i++)
                {
                    Receiver[i] = Receiver[i].Trim();
                }

                for (int index = 0; index < Receiver.Length; index++)
                {
                    using (MailMessage mm = new MailMessage(Sender, Receiver[index]))
                    {
                        DateTime now = DateTime.Now;
                        //DateTime realDate = now.AddMonths(-1);
                        DateTime realDate = DateTime.Now;
                        logger.Error("Trying Send Email At : " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));
                        mm.Subject = "The Monthly Report for Cost Variation";
                        mm.Body = string.Format(body, ReceiverName[index], CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(realDate.Month), realDate.Year);

                        System.Net.Mail.Attachment attachment;
                        attachment = new System.Net.Mail.Attachment(path);
                        mm.Attachments.Add(attachment);
                        mm.IsBodyHtml = true;
                        SmtpClient smtp = new SmtpClient();
                        smtp.Host = _Utilities.GetAppSetting("MailHost");
                        smtp.Port = 25;
                        smtp.EnableSsl = false;
                        System.Net.NetworkCredential credentials = new System.Net.NetworkCredential();
                        credentials.UserName = SenderUsername;
                        credentials.Password = SenderPassword;
                        smtp.UseDefaultCredentials = true;
                        smtp.Credentials = credentials;
                        smtp.Send(mm);
                        logger.Error("Email Send Successfully To : " + Receiver[index] + " at : " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));
                    }
                }
            }
            catch (Exception exp)
            {
                logger.Error("Exception Error : " + exp.Message);

            }



        }
    }
}
