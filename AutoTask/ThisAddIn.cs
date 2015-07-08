using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace AutoTask
{
    using System.Windows.Forms;

    public partial class ThisAddIn
    {
        private Office.CommandBar menuBar;
        private Office.CommandBarPopup newMenuBar;
        private Office.CommandBarButton buttonOne;
        private string menuTag = "AUniqueTag";
        private Outlook.NameSpace outlookNamespace;
        private Outlook.MAPIFolder inbox;
        private Outlook.Items items;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            RemoveMenubar();
            AddMenuBar();

            outlookNamespace = this.Application.GetNamespace("MAPI");
            inbox = outlookNamespace.GetDefaultFolder(
            Microsoft.Office.Interop.Outlook.
            OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);
        }

        void items_ItemAdd(object Item)
        {
            string filter = "USED CARS";
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            if (Item != null)
            {
                if (mail.Body.Contains("Sunil"))
                {
                    Outlook.TaskItem oTask = this.Application.CreateItem(Outlook.OlItemType.olTaskItem);
                    oTask.Subject = "This is my task subject";
                    oTask.DueDate = Convert.ToDateTime("06/25/2011");
                    oTask.StartDate = Convert.ToDateTime("06/20/2011");
                    oTask.ReminderSet = true;
                    oTask.ReminderTime = Convert.ToDateTime("06/28/2006 02:40:00 PM");
                    oTask.Body = mail.Body;
                    oTask.SchedulePlusPriority = "High";
                    oTask.Status = Microsoft.Office.Interop.Outlook.OlTaskStatus.olTaskInProgress;
                    oTask.Save();
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void AddMenuBar()
        {
            try
            {
                menuBar = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;
                newMenuBar = (Office.CommandBarPopup)menuBar.Controls.Add(
                    Office.MsoControlType.msoControlPopup, missing,
                    missing, missing, false);
                if (newMenuBar != null)
                {
                    newMenuBar.Caption = "See New Icon";
                    newMenuBar.Tag = menuTag;
                    buttonOne = (Office.CommandBarButton)
                        newMenuBar.Controls.
                    Add(Office.MsoControlType.msoControlButton, System.
                        Type.Missing, System.Type.Missing, 1, true);
                    buttonOne.Style = Office.MsoButtonStyle.
                        msoButtonIconAndCaption;
                    buttonOne.Caption = "New Icon";
                    buttonOne.FaceId = 65;
                    buttonOne.Tag = "c123";
                    buttonOne.Picture = getImage();
                    newMenuBar.Visible = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private stdole.IPictureDisp getImage()
        {
            stdole.IPictureDisp tempImage = null;
            try
            {
                System.Drawing.Icon newIcon =
                    Properties.Resources.Icon1;

                ImageList newImageList = new ImageList();
                newImageList.Images.Add(newIcon);
                tempImage = ConvertImage.Convert(newImageList.Images[0]);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return tempImage;
        }

        private void RemoveMenubar()
        {
            // If the menu already exists, remove it. 
            try
            {
                Office.CommandBarPopup foundMenu = (Office.CommandBarPopup)
                    this.Application.ActiveExplorer().CommandBars.ActiveMenuBar.
                    FindControl(Office.MsoControlType.msoControlPopup,
                    System.Type.Missing, menuTag, true, true);
                if (foundMenu != null)
                {
                    foundMenu.Delete(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
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

    sealed public class ConvertImage : System.Windows.Forms.AxHost
    {
        private ConvertImage()
            : base(null)
        {
        }

        public static stdole.IPictureDisp Convert
            (System.Drawing.Image image)
        {
            return (stdole.IPictureDisp)System.
                Windows.Forms.AxHost
                .GetIPictureDispFromPicture(image);
        }
    }
}
