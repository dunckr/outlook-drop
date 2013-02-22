using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office;

using System.Diagnostics;

namespace outlook_drop
{
    public partial class ThisAddIn
    {
        BrowserControl control;
        Office.Tools.CustomTaskPane explorerPane;

        Outlook.Explorer currentExplorer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            currentExplorer = this.Application.ActiveExplorer();

            control = new BrowserControl();
            explorerPane = CustomTaskPanes.Add(this.control, "Drop");
            explorerPane.DockPosition = Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            explorerPane.Width = 600;
            explorerPane.Visible = true;         

            control.Init();
            control.Login_EventHandler += new EventHandler(control_AuthorisationComplete);

            outlook_drop.Globals.Ribbons.ExplorerRibbon.uploadButton.Click += new Office.Tools.Ribbon.RibbonControlEventHandler(uploadButton_Click);
            outlook_drop.Globals.Ribbons.ExplorerRibbon.shareButton.Click += new Office.Tools.Ribbon.RibbonControlEventHandler(shareButton_Click);
        }
        private void control_AuthorisationComplete(object sender, EventArgs e)
        {
            explorerPane.Visible = false;
        }

        private Outlook.MailItem GetMailItem()
        {
            try
            {
                if (currentExplorer.Selection.Count > 0)
                {
                    Object selObject = currentExplorer.Selection[1];
                    if (selObject is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem = (selObject as Outlook.MailItem);
                        return mailItem;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            return null;
        }

        void uploadButton_Click(object sender, Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            try
            {
                Outlook.MailItem mailItem = GetMailItem();

                if (mailItem.Attachments.Count > 0)
                {
                    Outlook.Attachments attachments = mailItem.Attachments;
                    Outlook.Attachment attach = attachments[1];

                    string savePath = Path.Combine(Environment.GetFolderPath(
                                Environment.SpecialFolder.LocalApplicationData),
                                attach.FileName);
                    attach.SaveAsFile(savePath);

                    control.Upload(attach.FileName, savePath);
                    attach.Delete();
                    File.Delete(savePath);

                    // TODO: add multiple attachments
                    //for (int i = attachments.Count; i > 0; i--)
                    //{
                    //    attach = attachments[i];
                    //    mailItem.Attachments.Remove(i);                      
                    //}
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }


        }
        void shareButton_Click(object sender, Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            Outlook.MailItem mailItem = GetMailItem();
            Outlook.MailItem replyMail = ((Outlook._MailItem)mailItem).Reply();
            replyMail.Body = "DropBox share: " + control.GetShareLink() + replyMail.Body;
            replyMail.Display(true);
        }
        

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
