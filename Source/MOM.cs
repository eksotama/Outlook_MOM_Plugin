using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.DirectoryServices.AccountManagement;
using System.Drawing.Drawing2D;
using Outlook = Microsoft.Office.Interop.Outlook;
using YARTE.UI.Buttons;
using System.Globalization;
using System.IO;


namespace CalScanner
{

    public partial class MOM_Form : Form
    {

        public static Dictionary<String, String> emailNameMappingDict = null;
        public static Microsoft.Office.Interop.Outlook.Application app = null;

        private static String allEmailsFromContactString = "All Emails With This Contact";
        private static String allMeetingWithContactString = "All Meetings With This Contact";

        List<Outlook.MailItem> filteredItems = new List<Microsoft.Office.Interop.Outlook.MailItem>();
        Dictionary<String, Outlook.ContactItem> contactDict = new Dictionary<string, Outlook.ContactItem>();

        public static List<string> autoCompleteList = new List<string>();
        public static Attchments_MOM_Form attchMOMList = null;
        private bool toggleEventCall = false;
        private bool selectionComplete = false;
        private String inviteeListString = null;
        public static Microsoft.Office.Interop.Outlook.AppointmentItem nextMeetingItem = null;
        public static String[] timeValues = {
                                            "00:00 AM",
                                            "00:30 AM",
                                            "01:00 AM",
                                            "01:30 AM",
                                            "02:00 AM",
                                            "02:30 AM",
                                            "03:00 AM",
                                            "03:30 AM",
                                            "04:00 AM",
                                            "04:30 AM",
                                            "05:00 AM",
                                            "05:30 AM",
                                            "06:00 AM",
                                            "06:30 AM",
                                            "07:00 AM",
                                            "07:30 AM",
                                            "08:00 AM",
                                            "08:30 AM",
                                            "09:00 AM",
                                            "09:30 AM",
                                            "10:00 AM",
                                            "10:30 AM",
                                            "11:00 AM",
                                            "11:30 AM",
                                            "12:00 PM",
                                            "12:30 PM",
                                            "01:00 PM",
                                            "01:30 PM",
                                            "02:00 PM",
                                            "02:30 PM",
                                            "03:00 PM",
                                            "03:30 PM",
                                            "04:00 PM",
                                            "04:30 PM",
                                            "05:00 PM",
                                            "05:30 PM",
                                            "06:00 PM",
                                            "06:30 PM",
                                            "07:00 PM",
                                            "07:30 PM",
                                            "08:00 PM",
                                            "08:30 PM",
                                            "09:00 PM",
                                            "09:30 PM",
                                            "10:00 PM",
                                            "10:30 PM",
                                            "11:00 PM",
                                            "11:30 PM"};
        /// <summary>
        /// This method populates the attendee list in the attendee listbox
        /// </summary>
        /// <param name="inviteeList"></param>
        public void populateAttendeeList(String inviteeList)
        {
            Outlook.MailItem item = new Outlook.MailItem();

        }
        public MOM_Form(String calendarEntryId, Microsoft.Office.Interop.Outlook.MAPIFolder calendar, Microsoft.Office.Interop.Outlook._NameSpace ns, String inviteeList, Dictionary<String, String> mapDict, Microsoft.Office.Interop.Outlook.Application appl)
        {

            InitializeComponent();
            panel1.Height = 30;
            panel_next_meeting_atch1.Height = 25;
            btnMenuGroup1.Image = MOMOutlookAddIn.Properties.Resources.down;
            //This method will initialize the button toolbar for the text editor
            PredefinedButtonSets.SetupDefaultButtons(this.meetingNotes);
            //PredefinedButtonSets.SetupDefaultButtons(this.nextMeetingBody);
            //meetingNotes.ShowSelectionMargin = true;
            if (mapDict != null)
                emailNameMappingDict = mapDict;
            else
                emailNameMappingDict = new Dictionary<string, string>();
            app = appl;
            populateTimeCombo();
            button_Add_Action.Enabled = false;
            Outlook.AppointmentItem calItem = (Outlook.AppointmentItem)ns.GetItemFromID(calendarEntryId, calendar.StoreID);

            if (calItem != null)
            {
                meetingname.Text = calItem.Subject.Trim();
                meetingDate.Text = calItem.Start.ToString().Substring(0, calItem.Start.ToString().IndexOf(" "));
                starttime.Text = calItem.Start.ToString().Substring(calItem.Start.ToString().IndexOf(" ") + 1);
                endtime.Text = calItem.End.ToString().Substring(calItem.End.ToString().IndexOf(" ") + 1);
                location.Text = calItem.Location != null ? calItem.Location.Trim() : "";
                minutestaken.Text = System.DirectoryServices.AccountManagement.UserPrincipal.Current.DisplayName;
                chair.Text = calItem.Organizer.Trim();
                inviteeListString = inviteeList;
                nextMeetingItem = findNextOccuranceOfThisMeeting(meetingname.Text, calendar, calItem, inviteeList);
                populateAssignedToListBoxAndAttendeeCheckBox();
            }
            //minutestaken.Text=calItem.c

        }

        //This method reads local outlook data and extracts the details of people with whom the user has any calendar related interaction


        public void populateTimeCombo()
        {
            foreach (String t in timeValues)
            {
                nextmeetingStartTime.Items.Add(t);
                nextmeetingEndTime.Items.Add(t);
                comboBox_reminder.Items.Add(t);
            }
        }
        public void styleInviteeList(String inviteeList)
        {
            //This flag will stop uneccary call to the textbox texxtchanged evnet
            toggleEventCall = false;

            String[] inviteeArray = inviteeList.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            //textBox_next_invitees.Text = "";
            //textBox_next_invitees.au
            AutoCompleteStringCollection contactList = new AutoCompleteStringCollection();

            for (int i = 0; i < inviteeArray.Length; i++)
            {
                textBox_next_invitees.Text += inviteeArray[i];
                if (textBox_next_invitees.TextLength != 0) textBox_next_invitees.Text += "; ";
            }

            int start = 0;
            for (int j = 0; j < textBox_next_invitees.Text.Length; j++)
            {
                if (textBox_next_invitees.Text.ToString().Substring(j, 1).Equals(";"))
                {
                    this.textBox_next_invitees.SelectionStart = start;
                    this.textBox_next_invitees.SelectionLength = j - start;
                    textBox_next_invitees.SelectionFont = new Font(textBox_next_invitees.SelectionFont, FontStyle.Underline);
                    textBox_next_invitees.SelectionBackColor = Color.FromArgb(0, 215, 228, 188);
                    start = j + 3;
                }

            }


        }
        /// <summary>
        /// This method populates the assigned to list box and also adds a combo box column to the datagridview
        /// Also it populates the checked list box containing the attendee list 
        /// </summary>
        protected void populateAssignedToListBoxAndAttendeeCheckBox()
        {
            String[] inviteeArray = inviteeListString.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            DataGridViewComboBoxColumn ColumnItem = new DataGridViewComboBoxColumn();
            ColumnItem.HeaderText = "Assigned To";

            DataTable data = new DataTable();
            data.Columns.Add(new DataColumn("Value", typeof(string)));
            data.Columns.Add(new DataColumn("Description", typeof(string)));

            foreach (String s in inviteeArray)
            {
                if (!s.Equals("") && !listBox_AssgTo.Items.Contains(s))
                {

                    listBox_AssgTo.Items.Add(s.Trim());
                    //String attnd = FilterForm.emailNameMappingDict.ContainsKey(s.Trim()) ? FilterForm.emailNameMappingDict[s.Trim()].Trim() : s.Trim();
                    checkedListBox_attendee.Items.Add(s.Trim());
                    data.Rows.Add(s.Trim(), s.Trim());
                }
            }

            ColumnItem.DataSource = data;
            ColumnItem.ValueMember = "Value";
            ColumnItem.DisplayMember = "Description";

            // ColumnItem. = assgndList;
            dataGridView_action_items.Columns.Add(ColumnItem);

        }
        public Microsoft.Office.Interop.Outlook.AppointmentItem findNextOccuranceOfThisMeeting(String subject, Microsoft.Office.Interop.Outlook.MAPIFolder calendar, Outlook.AppointmentItem calItem, String inviteeList)
        {

            Microsoft.Office.Interop.Outlook.Items oItems = (Microsoft.Office.Interop.Outlook.Items)calendar.Items;

            DateTime startDate = calItem.Start;

            oItems.Sort("[Start]", false);
            oItems.IncludeRecurrences = true;

            String StringToCheck = "";
            StringToCheck = "[Start] > " + "\'" + startDate.ToString().Substring(0, startDate.ToString().IndexOf(" ") + 1).Trim() + "\'"
                 + " AND [Subject] = '" + calItem.Subject.Trim() + "'";
            // StringToCheck = "[Start] > " + "\'" + startDate.AddDays(1).ToString().Substring(0, startDate.ToString().IndexOf(" ")) + "\'"
            //                            + " AND [Subject] = '" + calItem.Subject.Trim() + "'";

            Microsoft.Office.Interop.Outlook.Items restricted;

            restricted = oItems.Restrict(StringToCheck);
            restricted.Sort("[Start]", false);

            restricted.IncludeRecurrences = true;
            Microsoft.Office.Interop.Outlook.AppointmentItem oAppt = (Microsoft.Office.Interop.Outlook.AppointmentItem)restricted.GetFirst();
                        
            if (oAppt != null) //Next occurance found
            {
                //Outlook.MailItem em= 
                Outlook.MailItem em = app.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
               
                em.HTMLBody = oAppt.Body;
                label_next_meeting.Visible = false;
                nextmeetingSubject.Text = oAppt.Subject.Trim();
                nextMeetingLocation.Text = oAppt.Location != null ? oAppt.Location.Trim() : "";
                //MemoryStream stream = new MemoryStream(oAppt.RTFBody);
                nextMeetingBody.Rtf = System.Text.Encoding.ASCII.GetString(oAppt.RTFBody);
                //ASCIIEncoding.Default.GetBytes(
                //byte[] b =
               // nextMeetingBody.Rtf=stream.
               // nextMeetingBody.Html = em.HTMLBody;               
                String timeTemp = oAppt.Start.ToString().Substring(oAppt.Start.ToString().IndexOf(" ") + 1);
                nextmeetingStartTime.Text = timeTemp.Substring(0, timeTemp.LastIndexOf(":")) + " " + (timeTemp.EndsWith("AM") ? "AM" : "PM");
                timeTemp = oAppt.End.ToString().Substring(oAppt.End.ToString().IndexOf(" ") + 1);
                nextmeetingEndTime.Text = timeTemp.Substring(0, timeTemp.LastIndexOf(":")) + " " + (timeTemp.EndsWith("AM") ? "AM" : "PM");
                dateTimePicker_nextMeetingDate.Checked = true;
                dateTimePicker_nextMeetingDate.Text = oAppt.Start.ToString();
                styleInviteeList(inviteeList);

                //Load the attachment list if there are already uploaded attachment details for the next meeting
                if (oAppt.Attachments != null && oAppt.Attachments.Count > 0)
                {
                    foreach (Outlook.Attachment atchItem in oAppt.Attachments)
                    {
                        dataGridView_next_meeting_atch.Rows.Add(atchItem.FileName.Trim(), "");
                    }
                }

                //textBox_next_invitees.Text = inviteeList;
                //=Convert.ToDateTime(oAppt.Start.ToString().Substring(oAppt.Start.ToString().IndexOf(" ") + 1));               
                return oAppt;
            }
            else
            {
                label_next_meeting.Visible = true;
                nextmeetingSubject.Text = subject;
                styleInviteeList(inviteeList);
                //textBox_next_invitees.Text = inviteeList;
                //textBox_next_invitees.textfr
                return null;
            }

        }
        private void hideAutoCompleteMenu()
        {
            listBox1.Visible = false;
        }
        private void listBox1_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (selectionComplete)
            {
                Rectangle rc = listBox1.GetItemRectangle(listBox1.SelectedIndex);
                LinearGradientBrush brush = new LinearGradientBrush(
                    rc, Color.Transparent, Color.Red, LinearGradientMode.ForwardDiagonal);
                Graphics g = Graphics.FromHwnd(listBox1.Handle);

                g.FillRectangle(brush, rc);

                if (listBox1.SelectedIndex >= 0)
                {
                    int index = listBox1.SelectedIndex;
                    String newItem = listBox1.Items[index].ToString();
                    textBox_next_invitees.Text = textBox_next_invitees.Text.ToString().Remove(textBox_next_invitees.Text.ToString().LastIndexOf(";") + 1);
                    int start = textBox_next_invitees.Text.Length + 2;
                    textBox_next_invitees.Text = textBox_next_invitees.Text + " " + newItem + "; ";
                    //styleInviteeList(textBox_next_invitees.Text);
                    /*this.textBox_next_invitees.SelectionStart = start;
                    this.textBox_next_invitees.SelectionLength = listBox1.Items[index].ToString().Length;
                    textBox_next_invitees.SelectionFont = new Font(textBox_next_invitees.SelectionFont, FontStyle.Underline);
                    textBox_next_invitees.SelectionBackColor = Color.FromArgb(0, 215, 228, 188);
                    textBox_next_invitees.Invalidate();*/
                    //styleInviteeList(textBox_next_invitees.Text.Trim());
                    hideAutoCompleteMenu();
                    selectionComplete = false;
                }
            }
        }
        private void textBox_next_invitees_KeyPress(object sender, KeyPressEventArgs e)
        {
            toggleEventCall = true;
            /*string keyword = "<";
            int count = 0;

            keyword += e.KeyChar;
            count++;
            Point point = this.textBox_next_invitees.GetPositionFromCharIndex(textBox_next_invitees.SelectionStart);
            point.Y += (int)Math.Ceiling(this.textBox_next_invitees.Font.GetHeight()) + 13; //13 is the .y postion of the richtectbox
            point.X += 105; //105 is the .x postion of the richtectbox*/
            /*listBox1.Location = point;
            listBox1.Show();
            listBox1.SelectedIndex = 0;
            listBox1.SelectedIndex = listBox1.FindString(keyword);*/
            // textBox_next_invitees.Focus();*/
        }
        private String getLatestString()
        {
            return textBox_next_invitees.Text.ToString().Substring(textBox_next_invitees.Text.ToString().LastIndexOf(";") + 1).Trim();
        }
        private void textBox_next_invitees_TextChanged(object sender, EventArgs e)
        {
            if (toggleEventCall)
            {
                listBox1.Items.Clear();
                if (textBox_next_invitees.Text.Length == 0)
                {
                    hideAutoCompleteMenu();
                    return;
                }


                String compareText = getLatestString();
                foreach (String s in autoCompleteList)
                {
                    if (compareText == null || compareText.Equals("") || s.StartsWith(compareText.Trim()))
                    {

                        // Point cursorPt = Cursor.Position;
                        //listBox1.Location = PointToClient(cursorPt);
                        listBox1.Items.Add(s);
                        //listBox1.Visible = true;                       

                    }
                }

                if (listBox1.Items.Count > 0)
                {
                    Point point = this.textBox_next_invitees.GetPositionFromCharIndex(textBox_next_invitees.SelectionStart);
                    point.Y += (int)Math.Ceiling(this.textBox_next_invitees.Font.GetHeight()) + 1;
                    point.X += 2;
                    listBox1.Location = point;
                    this.listBox1.BringToFront();
                    this.listBox1.Show();
                }
                //if (listBox1.Visible)
                // listBox1.Select();
            }
        }

        private void listBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                selectionComplete = true;
                textBox_next_invitees.Focus();
                listBox1_SelectedIndexChanged(sender, e);
                //Now set the cursor at the end of the line inside the textbox
                textBox_next_invitees.SelectionStart = textBox_next_invitees.Text.Length + 1;
                textBox_next_invitees.SelectionLength = 0;
            }
            if (e.KeyCode == Keys.Escape || e.KeyCode == Keys.Right || e.KeyCode == Keys.Left || e.KeyCode == Keys.Back)
            {
                listBox1.Visible = false;
                textBox_next_invitees.Focus();
            }
        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            selectionComplete = true;
            textBox_next_invitees.Focus();
            listBox1_SelectedIndexChanged(sender, e);
            //Now set the cursor at the end of the line inside the textbox
            textBox_next_invitees.SelectionStart = textBox_next_invitees.Text.Length + 1;
            textBox_next_invitees.SelectionLength = 0;
        }

        private void ShowTooltip(object img, EventArgs e)
        {
            if (img is PictureBox)
            {
                var imgBox = img as PictureBox;
                if (imgBox != null)
                {
                    toolTip_Info.SetToolTip(imgBox, imgBox.Tag.ToString());
                }
            }
            if (img is Label)
            {
                var imgBox = img as Label;
                if (imgBox != null)
                {
                    toolTip_Info.SetToolTip(imgBox, imgBox.Tag.ToString());
                }
            }

        }

        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button_Add_Action_Click(object sender, EventArgs e)
        {
            String assgndList = "";

            foreach (String item in listBox_AssgTo.SelectedItems)
            {
                if (!item.Equals(""))
                    assgndList = assgndList.Length > 0 ? assgndList + "; " + item.Trim() : item.Trim();
            }

            dataGridView_action_items.Rows.Add(
                richTextBox_Action_Item_Decsr.Text.Trim(),
                DateTime.Parse(dateTimePicker_Action_Item.Text).ToShortDateString(),
                !label_selected_file.Text.Trim().Equals("") ? label_selected_file.Text.ToString().Substring(label_selected_file.Text.ToString().LastIndexOf("\\") + 1) : "",
                !label_selected_file.Text.Trim().Equals("") ? label_selected_file.Text.ToString() : "",
                 assgndList);

            listBox_AssgTo.SelectedItems.Clear();
            label_selected_file.Text = "";
            button_Add_Action.Enabled = false;
        }

        private void button_Submit_Click(object sender, EventArgs e)
        {
            bool attendeeCheckPassed = false;

            if (checkedListBox_attendee.SelectedItems.Count == 0)
            {
                label_attendees.ForeColor = System.Drawing.Color.Red;
                DialogResult attndDialg = MessageBox.Show("You have not selected any attendee... Are you sure?", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (attndDialg == DialogResult.OK)
                    attendeeCheckPassed = true;
            }
            else
                attendeeCheckPassed = true;

            if (attendeeCheckPassed)
            {
                //First check and create the next meeting details

                if (nextMeetingItem == null)//Next meeting item not found - need to create one
                {
                    DialogResult nextMeetingConfirmation;
                    if (nextmeetingEndTime.Text.Trim().Equals(""))
                    {
                        nextMeetingConfirmation = MessageBox.Show("Do you want to create the next meeting?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                        if (nextMeetingConfirmation == DialogResult.Yes)
                        {
                            nextmeetingSubject.Focus();
                            return;
                        }
                    }
                    else
                    {
                        #region nextmeetingdetails
                        Outlook.AppointmentItem aptItem = (Microsoft.Office.Interop.Outlook.AppointmentItem)app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
                        aptItem.Subject = nextmeetingSubject.Text.Trim();
                        aptItem.Start = Convert.ToDateTime(dateTimePicker_nextMeetingDate.Text + " " + nextmeetingStartTime.Text.Trim());
                        aptItem.End = Convert.ToDateTime(dateTimePicker_nextMeetingDate.Text + " " + nextmeetingEndTime.Text.Trim());
                        aptItem.Location = nextMeetingLocation.Text;

                        aptItem.RTFBody = Encoding.ASCII.GetBytes(nextMeetingBody.Rtf);

                        if (dataGridView_next_meeting_atch.Rows.Count > 0) //Add attachments
                        {
                            foreach (DataGridViewRow dr in dataGridView_next_meeting_atch.Rows)
                            {
                                aptItem.Attachments.Add(dr.Cells[1].Value.ToString(), Outlook.OlAttachmentType.olByValue, 1, dr.Cells[0].Value.ToString());
                            }
                        }

                        String[] inviteeFinalList = textBox_next_invitees.Text.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                        String tempRecpAddressString="";
                        foreach (String s in inviteeFinalList)
                        {
                            if (s != null && !s.Trim().Equals(""))
                            {
                                  tempRecpAddressString=emailNameMappingDict.ContainsKey(s.Trim()) ? emailNameMappingDict[s.Trim()] : "";
                                  if (!tempRecpAddressString.Equals(""))
                                  {
                                      Outlook.Recipient recipient = aptItem.Recipients.Add(tempRecpAddressString);
                                      recipient.Type =
                              (int)Outlook.OlMeetingRecipientType.olRequired;
                                  }
                            }
                        }
                        aptItem.ReminderSet = true;
                        aptItem.ReminderMinutesBeforeStart = 15;
                        aptItem.BusyStatus = Microsoft.Office.Interop.Outlook.OlBusyStatus.olBusy;
                        aptItem.MeetingStatus = Microsoft.Office.Interop.Outlook.OlMeetingStatus.olMeeting;
                        aptItem.Save();
                        try
                        {
                            aptItem.Send();
                            label_next_meeting.Visible = false;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error sending invitation to recepients, one or more email address might be incorrect");
                        }
                        #endregion nextmeetingdetails
                    }
                    //((Outlook.AppointmentItem)aptItem).Send();

                }
                else //Next meeting item found - need to modify the details
                {
                    //nextMeetingItem                   
                    
                    nextMeetingItem.RTFBody = Encoding.ASCII.GetBytes(nextMeetingBody.Rtf);

                    nextMeetingItem.Start = Convert.ToDateTime(dateTimePicker_nextMeetingDate.Text + " " + nextmeetingStartTime.Text.Trim());
                    nextMeetingItem.End = Convert.ToDateTime(dateTimePicker_nextMeetingDate.Text + " " + nextmeetingEndTime.Text.Trim());
                    nextMeetingItem.Location = nextMeetingLocation.Text;

                    String[] inviteeFinalList = textBox_next_invitees.Text.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                    String tempRecpAddressString = "";
                    foreach (String s in inviteeFinalList)
                    {
                        if (s != null && !s.Trim().Equals(""))
                        {
                            tempRecpAddressString=emailNameMappingDict.ContainsKey(s.Trim()) ? emailNameMappingDict[s.Trim()] : "";
                            if (!tempRecpAddressString.Equals(""))
                            {
                                Outlook.Recipient recipient = nextMeetingItem.Recipients.Add(tempRecpAddressString);
                                recipient.Type =
                        (int)Outlook.OlMeetingRecipientType.olRequired;
                            }
                        }
                    }

                    if (dataGridView_next_meeting_atch.Rows.Count > 0) //Add attachments
                    {
                        foreach (DataGridViewRow dr in dataGridView_next_meeting_atch.Rows)
                        {
                            if (!dr.Cells[1].Value.ToString().Equals("")) //Full path empty - means this attachment was added using this window
                            nextMeetingItem.Attachments.Add(dr.Cells[1].Value.ToString(), Outlook.OlAttachmentType.olByValue, 1, dr.Cells[0].Value.ToString());
                        }
                    }

                    nextMeetingItem.Save();
                    try
                    {
                        nextMeetingItem.Send();
                        label_next_meeting.Visible = false;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error sending invitation to recepients, one or more email address might be incorrect");
                    }
                }
                //Assign tasks to participants
                String taskHTML = assignTasks();
                //Create the MOM HTML
                String MOMHTML = createMOMHtml(taskHTML);

                try
                {
                    Outlook.MailItem mailObj = app.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                    mailObj.To = convertRecpNamestoAddress(inviteeListString.Trim());
                    mailObj.HTMLBody = MOMHTML + mailObj.HTMLBody;
                    mailObj.Subject = "MOM " + meetingname.Text.Trim() + " - " + DateTime.Parse(meetingDate.Text.Trim()).ToShortDateString();


                    //Attach items in the email
                    foreach (DataGridViewRow dr in dataGridView_action_items.Rows)
                    {
                        if (!dr.Cells[3].Value.ToString().Equals("")) //Attachment in action items
                        {
                            String attachmentPath = dr.Cells[3].Value.ToString();
                            mailObj.Attachments.Add(dr.Cells[3].Value.ToString(), Outlook.OlAttachmentType.olByValue, 1, attachmentPath.Substring(attachmentPath.LastIndexOf("\\") + 1));
                        }
                    }

                    foreach (DataGridViewRow dr in dataGridView_attachments.Rows)
                    {
                        if (!dr.Cells[0].Value.ToString().Equals(""))
                        {
                            mailObj.Attachments.Add(dr.Cells[1].Value.ToString(), Outlook.OlAttachmentType.olByValue,
                                1, dr.Cells[0].Value.ToString());
                        }
                    }

                    mailObj.Send();
                    MessageBox.Show("Minutes sent successfuly", "", MessageBoxButtons.OK);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error sending MOM", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// This method goes through all the names passed in the input and converts these to email addresses
        /// </summary>
        /// <param name="invitelist"></param>
        /// <returns></returns>
        private static String convertRecpNamestoAddress(String invitelist)
        {
            String[] inviteeFinalList = invitelist.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            String returnString="";
            foreach (String s in inviteeFinalList)
               //   returnString = emailNameMappingDict.ContainsKey(s.Trim()) ? (returnString.Equals("") ? emailNameMappingDict[s.Trim()] : returnString + ";" + emailNameMappingDict[s.Trim()]) : returnString;
                returnString = emailNameMappingDict.ContainsKey(s.Trim()) ?
                    (returnString.Equals("") ? emailNameMappingDict[s.Trim()] : returnString + ";" + emailNameMappingDict[s.Trim()]) 
                    : (returnString.Equals("") ?s.Trim():returnString + ";"+s.Trim());

            return returnString;
        }

        /// <summary>
        /// This method creates the tasks, assigns and generates the html content to be attached in the assigned tasks section of the MOM
        /// </summary>
        /// <returns></returns>
        private String assignTasks()
        {
            Dictionary<int, String> taskList = new Dictionary<int, string>();
            String returnHTML = "";
            foreach (DataGridViewRow dr in dataGridView_action_items.Rows)
            {
                try
                {
                    //Use the Outlook application object created in the previous form
                    Outlook.TaskItem taskObj = app.CreateItem(Outlook.OlItemType.olTaskItem) as Outlook.TaskItem;

                    taskObj.Subject = dr.Cells[0].Value.ToString();
                    taskObj.Body = dr.Cells[0].Value.ToString();
                    taskObj.DueDate = DateTime.Parse(dr.Cells[1].Value.ToString());
                    //taskObj.Owner = dr.Cells[4].Value.ToString();
                    taskObj.Assign();
                    taskObj.Recipients.Add(dr.Cells[4].Value.ToString());
                    taskObj.ReminderSet = true;
                    taskObj.ReminderTime = Convert.ToDateTime(taskObj.DueDate.ToShortDateString() + " " + comboBox_reminder.Text.Trim());
                    if (!dr.Cells[3].Value.ToString().Equals(""))
                        taskObj.Attachments.Add(dr.Cells[3].Value.ToString(), Outlook.OlAttachmentType.olByValue, 1, dr.Cells[3].Value.ToString());
                    taskObj.Send();

                    //Create the HTML content for the MOM
                    returnHTML += "<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'>" +
                          "<td style='padding:.75pt .75pt .75pt .75pt'>" +
                          "<p class=MsoNormal style='margin-top:3.0pt;margin-right:0cm;margin-bottom:3.0pt;margin-left:0cm'><b><span style='font-size:8.0pt;font-family:\"Book Antiqua\",\"serif\"'>" +
                           "<a href=Outlook:" + taskObj.EntryID + ">" + taskObj.Subject + "</a>" +
                          "</span></b></p>  </td>" +
                          "<td style='padding:.75pt .75pt .75pt .75pt'>" +
                            "<p class=MsoNormal style='margin-top:3.0pt;margin-right:0cm;margin-bottom: 3.0pt;margin-left:0cm'><span style='font-size:8.0pt;font-family:\"Book Antiqua\",\"serif\"'>" +
                            taskObj.Owner +
                            "</span></b></p>  </td>" +
                             "<td style='padding:.75pt .75pt .75pt .75pt'>" +
                            "<p class=MsoNormal style='margin-top:3.0pt;margin-right:0cm;margin-bottom: 3.0pt;margin-left:0cm'><span style='font-size:8.0pt;font-family:\"Book Antiqua\",\"serif\"'>" +
                             taskObj.DueDate.ToShortDateString() +
                            "</span></b></p>  </td>" +
                            " </tr>";
                    //taskList.Add(dr.Index,taskObj.EntryID);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error Sending Task" + ex.ToString(), "Error", MessageBoxButtons.OK);
                }

            }
            return returnHTML;
            /*Outlook.MailItem mailObj = FilterForm.app.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mailObj.To = "shibasssmb@gmail.com";
            String htmlBody=mailObj.HTMLBody;
            foreach(KeyValuePair<int,String> kvp in taskList)
            {
                mailObj.HTMLBody = "<a href=Outlook:" + kvp.Value + ">" + kvp.Key.ToString() + "</a>";
            }*/
            //mailObj.HTMLBody += htmlBody;
            //mailObj.Send();
        }

        private String createMOMHtml(string taskHTML)
        {
            //var template = new HtmlTemplate(@"C:\Users\shibathethinker\Desktop\template.html");

            var template = new HtmlTemplate();
            String attendeeList = "";

            for (int i = 0; i < checkedListBox_attendee.CheckedIndices.Count; i++)
                attendeeList = attendeeList.Equals("") ? checkedListBox_attendee.Items[i].ToString() : attendeeList + "; " + checkedListBox_attendee.Items[i].ToString();

            var output = template.Render(new
            {
                TITLE = meetingname.Text.Trim(),
                DATEOFMEETING = meetingDate.Text.Trim(),
                STARTTIME = starttime.Text.Trim(),
                LOCATION = location.Text.Trim(),
                CHAIR = chair.Text.Trim(),
                MINUTETAKEN = minutestaken.Text.Trim(),
                ENDTIME = endtime.Text.Trim(),
                SUBJECT = meetingname.Text.Trim(),
                TOPIC = meetingNotes.Html,
                NEXTMEETINGDATE = dateTimePicker_nextMeetingDate.Text,
                NEXTMEETINGTIME = nextmeetingStartTime.Text,
                NEXTMEETINGLOCATION = nextMeetingLocation.Text,
                NEXTMEETINGSUBJECT = nextmeetingSubject.Text,
                ACTIONITEMS = taskHTML,
                ALLATTENDEES = attendeeList
                //METAKEYWORDS = "Keyword1, Keyword2, Keyword3",
                //BODY = "Body content goes here",
                //ETC = "etc"
            });
           // System.IO.File.WriteAllText(@"C:\Users\shibathethinker\Desktop\WriteLines.html", output);
            return output;
        }

        private void button_attach_file_actionItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileAttach = new OpenFileDialog();
            fileAttach.Multiselect = false;
            if (fileAttach.ShowDialog() == DialogResult.OK)
            {
                label_selected_file.Text = fileAttach.FileName;
                label_selected_file.Tag = fileAttach.FileName;

                //label_selected_file.
            }
        }

        private void delete_row_action_items(object sender, EventArgs e)
        {
            foreach (DataGridViewRow item in this.dataGridView_action_items.SelectedRows)
            {
                dataGridView_action_items.Rows.RemoveAt(item.Index);
            }
        }

        private void richTextBox_Action_Item_Decsr_TextChanged(object sender, EventArgs e)
        {
            if (richTextBox_Action_Item_Decsr.Text.Length > 0)
                button_Add_Action.Enabled = true;
            else
                button_Add_Action.Enabled = false;
        }

        private void textBox_next_invitees_KeyDown(object sender, KeyEventArgs e)
        {
            if (listBox1.Visible && (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down))
            {
                listBox1.Select();
                listBox1.SetSelected(0, true);
            }
        }

        private void buttonAttachFile_Click(object sender, EventArgs e)
        {

            OpenFileDialog fileAttach = new OpenFileDialog();
            fileAttach.Multiselect = false;
            if (fileAttach.ShowDialog() == DialogResult.OK)
            {
                //this.timer1.Enabled = true;
                String fullName = fileAttach.FileName;

                if (panel1.Height == 30)
                {
                    panel1.Height = (30 * 6) + 2;
                    btnMenuGroup1.Image = MOMOutlookAddIn.Properties.Resources.up;
                }
                /*else
                {
                    panel1.Height = 30;
                    btnMenuGroup1.Image = Properties.Resources.down;
                }*/
                dataGridView_attachments.Rows.Add(fullName.Substring(fullName.LastIndexOf("\\") + 1).Trim(), fullName.Trim());

                /*if (attchMOMList==null)
                  attchMOMList = new Attchments_MOM_Form();
                
                attchMOMList.loadData(fileAttach.FileName);
                attchMOMList.Location = PointToScreen(new Point(200, 200));
                attchMOMList.Show();
                attchMOMList.Activate();*/
                //label_selected_file.
            }
        }


        private void btnMenuGroup1_Click(object sender, EventArgs e)
        {
            if (panel1.Height == 30)
            {
                panel1.Height = (30 * 6) + 2;
                btnMenuGroup1.Image = MOMOutlookAddIn.Properties.Resources.up;
            }
            else
            {
                panel1.Height = 30;
                btnMenuGroup1.Image = MOMOutlookAddIn.Properties.Resources.down;
            }
        }

        private void toolStripMenuItem_MOM_delete_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow item in this.dataGridView_attachments.SelectedRows)
            {
                dataGridView_attachments.Rows.RemoveAt(item.Index);
            }
        }

        private void buttonAttachFile_NextMeeting_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileAttach = new OpenFileDialog();
            fileAttach.Multiselect = false;
            if (fileAttach.ShowDialog() == DialogResult.OK)
            {
                //this.timer1.Enabled = true;
                String fullName = fileAttach.FileName;

                if (panel_next_meeting_atch1.Height == 25)
                {
                    panel_next_meeting_atch1.Height = (25 * 6) + 2;
                    button_next_meeting_atch1.Image = MOMOutlookAddIn.Properties.Resources.up;
                }

                dataGridView_next_meeting_atch.Rows.Add(fullName.Substring(fullName.LastIndexOf("\\") + 1).Trim(), fullName.Trim());
            }






        }

        private void button_next_meeting_atch1_Click(object sender, EventArgs e)
        {
            if (panel_next_meeting_atch1.Height == 25)
            {
                panel_next_meeting_atch1.Height = (25 * 6) + 2;
                button_next_meeting_atch1.Image = MOMOutlookAddIn.Properties.Resources.up;
            }
            else
            {
                panel_next_meeting_atch1.Height = 25;
                button_next_meeting_atch1.Image = MOMOutlookAddIn.Properties.Resources.down;
            }
        }

        private void toolStripMenuItemDeleteNextMeetingAtch_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow item in this.dataGridView_next_meeting_atch.SelectedRows)
            {
                dataGridView_next_meeting_atch.Rows.RemoveAt(item.Index);
            }
        }


        private void nextmeetingStartTime_SelectedIndexChanged(object sender, EventArgs e)
        {
            String currentValue = nextmeetingStartTime.Text;
            nextmeetingEndTime.Text = DateTime.ParseExact(currentValue, "hh:mm tt", CultureInfo.CurrentCulture).AddMinutes(30).ToShortTimeString();

        }

        private void expand_collapse_Click(object sender, EventArgs e)
        {
            if (nextMeetingBody.Height == 168)
                nextMeetingBody.Height = 300;
            else
                nextMeetingBody.Height = 168;
        }

        private void nextMeetingBody_Load(object sender, EventArgs e)
        {

        }

        private void nextMeetingBody_Enter(object sender, EventArgs e)
        {
            nextMeetingBody.Height = 300;
        }

        private void nextMeetingBody_Leave(object sender, EventArgs e)
        {
            nextMeetingBody.Height = 187;
        }

        private void meetingNotes_Enter(object sender, EventArgs e)
        {
            meetingNotes.Height = 350;
            splitContainer1.Height = 400;
        }

        private void meetingNotes_Leave(object sender, EventArgs e)
        {
            meetingNotes.Height = 270;
            splitContainer1.Height = 360; 
        }
    }
}
