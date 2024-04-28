using CsvHelper;
using CsvHelper.Configuration.Attributes;
using System.Globalization;
using System.Runtime.InteropServices;
using WindowsInput;
using WindowsInput.Native;
using static CSV_Viewer.Form1;

namespace CSV_Viewer
{
    public partial class Form1 : Form
    {

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private InputSimulator inputSimulator = new InputSimulator();

        const int WM_HOTKEY = 0x0312;
        const int HOTKEY_ID = 1; // Unique identifier for your hotkey
        const uint MOD_CONTROL = 0x02; // Modifier key
        const uint MOD_SHIFT = 0x04; // Modifier key
        const uint VK_BACKSLASH = 0xDC; // Virtual key code for '\'

        private Label labelStatus = new Label();
        private List<DonationRecord> records = new List<DonationRecord>();
        private int currentIndex = 0;
        private TabControl tabControl = new TabControl();
        private TabPage tabPP = new TabPage("PP");
        private TabPage tabCC = new TabPage("CC");
        private ListView listViewPP = new ListView();
        private ListView listViewCC = new ListView();

        private Button buttonNextPP = new Button();
        private Button buttonPreviousPP = new Button();
        private Label labelStatusPP = new Label();

        private Button buttonNextCC = new Button();
        private Button buttonPreviousCC = new Button();
        private Label labelStatusCC = new Label();

        private int currentIndexPP = 0;
        private int currentIndexCC = 0;

        private List<DonationRecord> recordsPP = new List<DonationRecord>();
        private List<DonationRecord> recordsCC = new List<DonationRecord>();

        private ToolTip tooltip = new ToolTip();

        private bool hotkeyNavigation = false;


        public class DonationRecord
        {
            [Name("WEB_ID")]
            public string WebId { get; set; }

            [Name("Title")]
            public string Title { get; set; }

            [Name("Forename")]
            public string Forename { get; set; }

            [Name("Surname")]
            public string Surname { get; set; }

            [Name("Address_line1")]
            public string AddressLine1 { get; set; }

            [Name("Address_line2")]
            public string AddressLine2 { get; set; }

            [Name("Address_line3")]
            public string AddressLine3 { get; set; }

            [Name("Town")]
            public string Town { get; set; }

            [Name("County")]
            public string County { get; set; }

            [Name("Postcode")]
            public string Postcode { get; set; }

            [Name("Country")]
            public string Country { get; set; }

            [Name("Email_address")]
            public string EmailAddress { get; set; }

            [Name("Mobile_phone_number")]
            public string MobilePhoneNumber { get; set; }

            [Name("Landline_phone_number")]
            public string LandlinePhoneNumber { get; set; }

            [Name("Amount")]
            public string Amount { get; set; }

            [Name("In memory of")]
            public string InMemoryOf { get; set; }

            [Name("Name of loved one")]
            public string NameOfLovedOne { get; set; }

            [Name("In memory of other")]
            public string InMemoryOfOther { get; set; }

            [Name("Source")]
            public string Source { get; set; }

            [Name("Tribute Fund Number")]
            public string TributeFundNumber { get; set; }

            [Name("Gift_Aid")]
            public string GiftAid { get; set; }

            [Name("Gift_Aid_Confirmation_Date")]
            public string GiftAidConfirmationDate { get; set; }

            [Name("Mail_DPA")]
            public string MailDPA { get; set; }

            [Name("Email_DPA")]
            public string EmailDPA { get; set; }

            [Name("Telephone_DPA")]
            public string TelephoneDPA { get; set; }

            [Name("SMS_DPA")]
            public string SMSDPA { get; set; }

            [Name("Donation_Date")]
            public string DonationDate { get; set; }

            [Name("Datacash_reference")]
            public string DatacashReference { get; set; }

            [Name("DOB")]
            public string DOB { get; set; }

            [Name("Payment_method")]
            public string PaymentMethod { get; set; }

            public string DonorName => $"{Forename} {Surname}";
            public string InMemOfForenames => NameOfLovedOne?.Split(' ').SkipLast(1).Aggregate((current, next) => current + " " + next);
            public string InMemoryOfSurname => NameOfLovedOne?.Split(' ').LastOrDefault();
            public string DonorGender { get; private set; }  // Making it private set so it can only be modified internally within the class
            public string InMemFrom => DetermineInMemFrom(InMemoryOf, DonorGender);
            public string FromCode => DetermineCode(InMemoryOf);
            public string ToCode => DetermineCode(InMemFrom);

            public bool TributeNumberValid => (TributeFundNumber != null) ? TributeFundNumber.Length > 5 : false;

            private ToolTip tooltip = new ToolTip();

            public bool ManuallySelectedAsTributeFund = false;



            private string NormalizeRelationship(string relationship)
            {
                return relationship?.ToLower().Replace("-", " ").Replace("  ", " ").Trim();
            }

            public void UpdateGender(string newGender)
            {
                DonorGender = newGender;
            }


            public void SetGender(string title)
            {
                switch (title.ToLower().Trim())
                {
                    case "mr":
                        DonorGender = "Male";
                        break;
                    case "mrs":
                    case "miss":
                    case "ms":
                        DonorGender = "Female";
                        break;
                    default:
                        DonorGender = "Unknown";
                        break;
                }
            }

            private string DetermineInMemFrom(string inMemoryOf, string donorGender)
            {
                var normalizedMemoryOf = NormalizeRelationship(inMemoryOf);
                var inverseRelationships = new Dictionary<string, (string Male, string Female)>
    {
        {"grandfather", ("Granddaughter", "Grandson")},
        {"grandmother", ("Granddaughter", "Grandson")},
        {"father", ("Daughter", "Son")},
        {"mother", ("Daughter", "Son")},
        {"uncle", ("Niece", "Nephew")},
        {"aunt", ("Niece", "Nephew")},
        {"sister", ("Sister", "Brother")},
        {"brother", ("Sister", "Brother")},
        {"brother in law", ("Sister in law", "Brother in law")},
        {"sister in law", ("Sister in law", "Brother in law")},
        {"husband", ("Wife", "Wife")},
        {"wife", ("Husband", "Husband")},
        {"father in law", ("Daughter in law", "Son in law")},
        {"mother in law", ("Daughter in law", "Son in law")},
        {"son in law", ("Mother in law", "Father in law")},
        {"daughter in law", ("Mother in law", "Father in law")}
    };

                // Mutual relationships with specific descriptions
                var mutualRelationshipMappings = new Dictionary<string, string>
    {
        {"cousin", "Cousin"},
        {"friend", "Friend of"},
        {"neighbour", "Neighbour"},
        {"colleague", "Colleague of"},
        {"partner", "Partner of"},
        {"manager", "Manager of"},  // if "manager" is used instead of "manager of"
        {"business partner", "Business Partner"},
        {"joint fundraiser", "Joint Fundraiser"},  // Example adjustments
        // add other mappings as needed to match the relationship codes
    };

                if (mutualRelationshipMappings.TryGetValue(normalizedMemoryOf, out var formattedRelationship))
                {
                    return formattedRelationship; // Use the specified relationship format
                }

                if (inverseRelationships.TryGetValue(normalizedMemoryOf, out var opposite))
                {
                    return donorGender == "Male" ? opposite.Male : donorGender == "Female" ? opposite.Female : "Unknown";
                }

                return "Unknown"; // For relationships without a clear inverse or unhandled cases
            }

            private string DetermineCode(string relationship)
            {
                var normalizedRelationship = NormalizeRelationship(relationship);
                var relationshipCodes = new Dictionary<string, string>
        {
        {"aunt", "AUN"},
        {"board member of", "BM01"},
        {"brother", "BRO"},
        {"brother in law", "BIL"},
        {"business partner", "BP"},
        {"chairman of", "CH01"},
        {"child of", "CD01"},
        {"child/dependant", "PC04"},
        {"colleague of", "CO01"},
        {"collector", "CL01"},
        {"community fundraiser (speaker)", "SC04"},
        {"consultant to", "CON"},
        {"contributor to in memoriam", "IMCO"},
        {"cousin", "COU"},
        {"daughter", "DAU"},
        {"daughter in law", "DIL"},
        {"dependant of", "DO01"},
        {"deputy coordinator", "DC07"},
        {"district fundraiser", "DF19"},
        {"donate to", "DT01"},
        {"emergency contact details", "EM01"},
        {"event delegate guest", "EVD1"},
        {"executor", "EXEC"},
        {"father", "FTR"},
        {"father in law", "FIL"},
        {"fiance", "FIN1"},
        {"fiancee", "FIN2"},
        {"friend of", "FO01"},
        {"goddaughter", "GDR"},
        {"godfather", "GFR"},
        {"godmother", "GMR"},
        {"godson", "GDS"},
        {"granddaughter", "GDT"},
        {"grandfather", "GFT"},
        {"grandmother", "GMT"},
        {"grandson", "GSN"},
        {"great aunt", "GAU"},
        {"great nephew", "GNE"},
        {"great niece", "GNI"},
        {"great uncle", "GUN"},
        {"guide dog", "GDOG"},
        {"guide dog owner", "GDO"},
        {"guide runner for", "GR1"},
        {"husband", "HUS"},
        {"individual to individual", "I2I"},
        {"is agent for", "AGF"},
        {"is client of", "AG"},
        {"joint fundraiser", "JF03"},
        {"joint legator", "JLEG"},
        {"joint to individual", "I2J"},
        {"legacy engagement contact", "LE01"},
        {"manager of", "MO01"},
        {"married to", "MT01"},
        {"member", "MEM"},
        {"mgm recommendee", "MGRE"},
        {"mgm recommender", "MGRR"},
        {"mother", "MTR"},
        {"mother in law", "MIL"},
        {"neighbour", "NEI1"},
        {"nephew", "NEP"},
        {"niece", "NIE"},
        {"parent of", "PO01"},
        {"partner of", "PT01"},
        {"personal assistant", "PA"},
        {"raffle buyer", "RAFB"},
        {"recipient of in memoriam", "IMRP"},
        {"relative of", "RO01"},
        {"sister", "SIS"},
        {"sister in law", "SISL"},
        {"son", "SON"},
        {"son in law", "SIL"},
        {"spouse", "SPOU"},
        {"step daughter", "STD"},
        {"step father", "STF"},
        {"step mother", "STM"},
        {"step son", "STS"},
        {"subordinate of", "SB01"},
        {"supplier of", "SP01"},
        {"uncle", "UNC"},
        {"volunteer for", "VO02"},
        {"wife", "WIFE"},
            // Ensure these codes are correct as per your specific requirements
        };

                if (relationshipCodes.TryGetValue(normalizedRelationship, out var code))
                {
                    return code;
                }

                return "Unknown"; // For relationships without a code or unhandled cases
            }
        }


        private DateTime? ParseDonationDate(string dateString)
        {
            DateTime date;
            // Adjusting the format to parse both date and time
            if (DateTime.TryParseExact(dateString, "dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            {
                return date;
            }
            return null; // Or handle parsing failure accordingly
        }





        public Form1()
        {
            InitializeComponent();
            SetupForm();
            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, VK_BACKSLASH);  // ID 1 for Ctrl+|
            RegisterHotKey(this.Handle, HOTKEY_ID + 1, MOD_CONTROL | MOD_SHIFT, VK_BACKSLASH); // ID 2 for Ctrl+Shift+|
            InitializeTooltip();
            listViewPP.ItemSelectionChanged += ListView_ItemSelectionChanged;
            listViewCC.ItemSelectionChanged += ListView_ItemSelectionChanged;

        }

        private void InitializeTooltip()
        {
            tooltip.AutoPopDelay = 5000;
            tooltip.InitialDelay = 1000;
            tooltip.ReshowDelay = 500;
            tooltip.ShowAlways = true;
            tooltip.OwnerDraw = true;  // Enable owner draw

            tooltip.Draw -= tooltip_Draw;  // Remove existing event handler if any to avoid duplicates
            tooltip.Popup -= tooltip_Popup;

            tooltip.Draw += tooltip_Draw;  // Assign event handler for drawing
            tooltip.Popup += tooltip_Popup;  // Assign event handler for setting the size
        }

        private void tooltip_Draw(object sender, DrawToolTipEventArgs e)
        {
            using (Font f = new Font("Arial", 16, FontStyle.Bold))  // Ensure the font is disposed of properly
            {
                e.DrawBackground();
                e.DrawBorder();
                e.Graphics.DrawString(e.ToolTipText, f, Brushes.Black, new PointF(2, 2));
            }
        }

        private void tooltip_Popup(object sender, PopupEventArgs e)
        {
            using (Font f = new Font("Arial", 16, FontStyle.Bold))  // Example font configuration
            {
                e.ToolTipSize = TextRenderer.MeasureText(tooltip.GetToolTip(e.AssociatedControl), f);
            }
        }




        private void SetupForm()
        {

            // Setup for PP tab
            buttonNextPP.Text = "Next";
            buttonNextPP.Location = new System.Drawing.Point(300, 460); // Example position
            buttonNextPP.Size = new System.Drawing.Size(75, 23);
            tabPP.Controls.Add(buttonNextPP);

            buttonPreviousPP.Text = "Previous";
            buttonPreviousPP.Location = new System.Drawing.Point(220, 460); // Example position
            buttonPreviousPP.Size = new System.Drawing.Size(75, 23);
            tabPP.Controls.Add(buttonPreviousPP);

            labelStatusPP.Location = new System.Drawing.Point(10, 460); // Example position
            labelStatusPP.Size = new System.Drawing.Size(200, 23);
            labelStatusPP.Text = "Record 0 of 0";
            tabPP.Controls.Add(labelStatusPP);

            // Setup for CC tab (similarly adjust positions as per your layout)
            buttonNextCC.Text = "Next";
            buttonNextCC.Location = new System.Drawing.Point(300, 460); // Adjust as per layout
            buttonNextCC.Size = new System.Drawing.Size(75, 23);
            tabCC.Controls.Add(buttonNextCC);

            buttonPreviousCC.Text = "Previous";
            buttonPreviousCC.Location = new System.Drawing.Point(220, 460); // Adjust as per layout
            buttonPreviousCC.Size = new System.Drawing.Size(75, 23);
            tabCC.Controls.Add(buttonPreviousCC);

            labelStatusCC.Location = new System.Drawing.Point(10, 460); // Adjust as per layout
            labelStatusCC.Size = new System.Drawing.Size(200, 23);
            labelStatusCC.Text = "Record 0 of 0";
            tabCC.Controls.Add(labelStatusCC);

            buttonNextPP.Text = "Next PP";
            buttonNextPP.Click += (sender, e) => NavigateRecords("PP", true);
            tabPP.Controls.Add(buttonNextPP);

            buttonPreviousPP.Text = "Previous";
            buttonPreviousPP.Click += (sender, e) => NavigateRecords("PP", false);
            tabPP.Controls.Add(buttonPreviousPP);

            labelStatusPP.Text = "Record 0 of 0";
            tabPP.Controls.Add(labelStatusPP);

            buttonNextCC.Text = "Next CC";
            buttonNextCC.Click += (sender, e) => NavigateRecords("CC", true);
            tabPP.Controls.Add(buttonNextPP);

            buttonPreviousCC.Text = "Previous CC";
            buttonPreviousCC.Click += (sender, e) => NavigateRecords("CC", false);
            tabPP.Controls.Add(buttonPreviousPP);

            labelStatusPP.Text = "Record 0 of 0";
            tabPP.Controls.Add(labelStatusPP);

            // Initialize listViewPP for PP Tab
            listViewPP.View = View.Details;
            listViewPP.FullRowSelect = true;
            listViewPP.GridLines = true;
            listViewPP.Dock = DockStyle.Fill;
            listViewPP.Columns.Add("#", 25);
            listViewPP.Columns.Add("Type", 150);
            listViewPP.Columns.Add("Value", 190);
            listViewPP.HideSelection = false;
            listViewPP.MultiSelect = false;

            // Initialize listViewCC for CC Tab
            listViewCC.View = View.Details;
            listViewCC.FullRowSelect = true;
            listViewCC.GridLines = true;
            listViewCC.Dock = DockStyle.Fill;
            listViewCC.Columns.Add("#", 25);
            listViewCC.Columns.Add("Type", 150);
            listViewCC.Columns.Add("Value", 190);
            listViewCC.HideSelection = false;
            listViewCC.MultiSelect = false;

            // Setup TabControl
            tabPP.Controls.Add(listViewPP);
            tabCC.Controls.Add(listViewCC);
            tabControl.Controls.Add(tabCC);
            tabControl.Controls.Add(tabPP);
            tabControl.Location = new System.Drawing.Point(10, 30);
            tabControl.Size = new System.Drawing.Size(400, 520);
            this.Controls.Add(tabControl);

            // Setup buttons and label
            // Initialize MenuStrip
            MenuStrip menuStrip = new MenuStrip();
            ToolStripMenuItem fileMenuItem = new ToolStripMenuItem("File");
            ToolStripMenuItem openMenuItem = new ToolStripMenuItem("Open");
            ToolStripMenuItem aboutMenuItem = new ToolStripMenuItem("About");

            // Add the Open menu item under File
            fileMenuItem.DropDownItems.Add(openMenuItem);

            // Add the About menu item to the menu strip
            menuStrip.Items.Add(fileMenuItem);
            menuStrip.Items.Add(aboutMenuItem);

            // Add MenuStrip to the form
            this.Controls.Add(menuStrip);
            this.MainMenuStrip = menuStrip;

            // Event handler for Open
            openMenuItem.Click += new EventHandler(buttonLoadCsv_Click);

            // Event handler for About
            aboutMenuItem.Click += (sender, e) =>
            {
                MessageBox.Show("Made my Michael. Oh Yeah!", "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };


            labelStatus.Location = new System.Drawing.Point(10, 310);
            labelStatus.Size = new System.Drawing.Size(200, 20);
            this.Controls.Add(labelStatus);

            // Form Settings
            this.Text = "CSV Record Viewer";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new System.Drawing.Size(435, 600);

            tabControl.SelectedIndexChanged += TabControl_SelectedIndexChanged;

        }


        private bool isInBatchDetails = true; // Initial state

        private void SwitchDetails()
        {
            if (isInBatchDetails)
            {
                // Logic to switch to record details
                isInBatchDetails = false;
                // Example: Set focus to the first record detail
            }
            else
            {
                // Logic to switch back to batch details or to the next record
                isInBatchDetails = true;
                // Example: Set focus back to batch details or to the next record's batch details
            }

            // Example clipboard operation (adjust according to your "value" column)
            if (listViewPP.SelectedItems.Count > 0)
            {
                var item = listViewPP.SelectedItems[0];
                Clipboard.SetText(item.SubItems[2].Text); // Assuming the value is in the third column
            }
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == WM_HOTKEY)
            {
                if (m.WParam.ToInt32() == HOTKEY_ID)
                {
                    // Handle Ctrl+| for regular cycling
                    CycleThroughListView(false);
                }
                else if (m.WParam.ToInt32() == HOTKEY_ID + 1)
                {
                    // Handle Ctrl+Shift+| for cycling and pasting
                    CycleThroughListView(true);
                }
            }
        }

        private void CycleThroughListView(bool pasteAfterCycling)
        {
            ListView activeListView = GetActiveListView();
            if (activeListView == null) return;

            int nextIndex = GetNextIndex(activeListView);

            if (nextIndex != -1)
            {
                SelectAndCopyItem(activeListView, nextIndex);
                if (pasteAfterCycling)
                {
                    // Simulate pasting the content, which might involve setting text in a textbox or another component that accepts input
                    PasteClipboardContent();
                }
                var item = activeListView.Items[nextIndex];
                ResetAndShowTooltip(activeListView, $"Copied: {item.SubItems[1].Text} - {item.SubItems[2].Text}");

            }
            else
            {
                if (activeListView == listViewPP)
                {
                    buttonNextPP.PerformClick();
                }
                else if (activeListView == listViewCC)
                {
                    buttonNextCC.PerformClick();
                }
                if (activeListView.Items.Count > 0)
                {
                    nextIndex = GetNextIndex(activeListView);
                    SelectAndCopyItem(activeListView, nextIndex);
                    if (pasteAfterCycling)
                    {
                        PasteClipboardContent();
                    }
                }
            }
        }

        private void PasteClipboardContent()
        {
            inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);
        }

        private void SelectAndCopyItem(ListView activeListView, int nextIndex)
        {

            // Focus the next item
            activeListView.Focus();
            activeListView.Items[nextIndex].Selected = true;
            activeListView.Items[nextIndex].EnsureVisible();

            // Copy the item's value to the clipboard if it's not a section header
            if (!activeListView.Items[nextIndex].Text.StartsWith("Batch Data") && !activeListView.Items[nextIndex].Text.StartsWith("Record Data"))
            {
                string copiedValue = activeListView.Items[nextIndex].SubItems[2].Text;
                string copiedType = activeListView.Items[nextIndex].SubItems[1].Text;
                Clipboard.SetText(copiedValue);

                // Show the tooltip with copied type and value
                string tooltipText = $"Copied: {copiedType} - {copiedValue}";
                ResetAndShowTooltip(activeListView, tooltipText);

            }
        }

        private ListView GetActiveListView()
        {
            // Check which tab is currently selected and return the corresponding ListView
            if (tabControl.SelectedTab == tabPP)
            {
                return listViewPP; // Return the ListView associated with the PP tab
            }
            else if (tabControl.SelectedTab == tabCC)
            {
                return listViewCC; // Return the ListView associated with the CC tab
            }

            return null; // Return null if no matching tab is found (shouldn't happen if tabs are properly set up)
        }

        private ListViewItem lastSelectedItem = null; // Keep track of the last selected item

        private int GetNextIndex(ListView listView)
        {
            // Starting from the current position, look for the next index
            for (int i = currentIndex + 1; i < listView.Items.Count; i++)
            {
                if (!string.IsNullOrEmpty(listView.Items[i].SubItems[0].Text))
                {
                    currentIndex = i;
                    return i;
                }
            }

            return -1; // Return -1 when no more items are found
        }

        // Implement the button click event handler to reset to Batch Details
        private void goBackToBatchDetailsButton_Click(object sender, EventArgs e)
        {
            isInBatchDetails = true;
            currentIndex = -1; // Reset or set to the index of the first item in Batch Data
                               // Additional logic to focus the first item in Batch Data if needed
        }



        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Cleanup
            UnregisterHotKey(this.Handle, HOTKEY_ID);  // Unregister Ctrl+|
            UnregisterHotKey(this.Handle, HOTKEY_ID + 1); // Unregister Ctrl+Shift+|

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Intentionally left blank
        }

        private void buttonLoadCsv_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var filePath = openFileDialog.FileName;
                    LoadRecords(filePath);
                }
            }
        }

        private void NavigateRecords(string paymentType, bool next)
        {
            if (paymentType == "PP")
            {

                int newIndexPP = next ? Math.Min(currentIndexPP + 1, recordsPP.Count - 1) : Math.Max(currentIndexPP - 1, -1);

                if (newIndexPP == currentIndexPP)
                {
                    return;
                }

                currentIndexPP = newIndexPP;

                //reset cycle to copy
                currentIndex = 0;

                if (currentIndexPP >= 0)
                {
                    DisplayCurrentRecord("PP");
                }
                else
                {
                    DisplayBatchData("PP");
                }
            }
            else if (paymentType == "CC")
            {
                int newIndexCC = next ? Math.Min(currentIndexCC + 1, recordsCC.Count - 1) : Math.Max(currentIndexCC - 1, -1);

                if (newIndexCC == currentIndexCC)
                {
                    return;
                }

                currentIndexCC = newIndexCC;

                currentIndex = 0;

                if (currentIndexCC >= 0)
                {
                    DisplayCurrentRecord("CC");
                }
                else
                {
                    DisplayBatchData("CC");
                }

            }
        }

        private void DisplayCurrentRecord(string paymentType)
        {
            var activeRecords = paymentType == "PP" ? recordsPP : recordsCC;
            var currentIndex = paymentType == "PP" ? currentIndexPP : currentIndexCC;
            var activeListView = paymentType == "PP" ? listViewPP : listViewCC;
            var activeStatusLabel = paymentType == "PP" ? labelStatusPP : labelStatusCC;

            if (activeRecords.Count == 0)
            {
                activeStatusLabel.Text = "No records found";
                return;
            }

            if (currentIndex < 0 || currentIndex >= activeRecords.Count)
            {
                activeStatusLabel.Text = "Index out of range";
                return;
            }

            var recordToShow = activeRecords[currentIndex];

            // Append the current record's details to the ListView
            AddRecordDetailsToListView(activeListView, recordToShow);

            // Update the status label
            activeStatusLabel.Text = $"Record {currentIndex + 1} of {activeRecords.Count}";
        }

        private void DisplayBatchData(string paymentType)
        {
            var records = paymentType == "PP" ? recordsPP : recordsCC;
            var listView = paymentType == "PP" ? listViewPP : listViewCC;
            var statusLabel = paymentType == "PP" ? labelStatusPP : labelStatusCC;

            listView.Items.Clear(); // Clear the list view for new data

            if (records.Count == 0)
            {
                statusLabel.Text = "No batch data found";
                return;
            }

            // Calculate and display batch data
            var boldFont = new Font(listView.Font, FontStyle.Bold);
            listView.Items.Add(new ListViewItem(new[] { "", "Batch Detail", "" }) { BackColor = Color.LightGray, Font = boldFont });
            listView.Items.Add(new ListViewItem(new[] { "1", "Batch Type", "Bank statement processing" }) { BackColor = Color.Azure, Font = boldFont });
            listView.Items.Add(new ListViewItem(new[] { "2", "Bank Account", "008" }) { BackColor = Color.Azure, Font = boldFont });
            listView.Items.Add(new ListViewItem(new[] { "3", "Account Type", "P" }) { BackColor = Color.Azure, Font = boldFont });

            double totalAmount = records.Sum(record => double.TryParse(record.Amount, out double amt) ? amt : 0);
            listView.Items.Add(new ListViewItem(new[] { "4", "TOTAL", totalAmount.ToString("N2") }) { BackColor = Color.Azure, Font = boldFont });  // Format total amount with 2 decimal places

            // Update status label
            statusLabel.Text = "Batch details displayed";
        }

        private void AddRecordDetailsToListView(ListView listView, DonationRecord record)
        {
            listView.Items.Clear();
            var boldFont = new Font(Font, FontStyle.Bold);

            listView.Items.Add(new ListViewItem(new[] { "", "STEPS", "" }) { BackColor = Color.LightGray });
            listView.Items.Add(new ListViewItem(new[] { "1", "WebId", record.WebId }) { BackColor = Color.Azure, Font = boldFont });
            listView.Items.Add(new ListViewItem(new[] { "2", "DatacashReference", record.DatacashReference }) { BackColor = Color.Azure, Font = boldFont });
            listView.Items.Add(new ListViewItem(new[] { "3", "Source", record.Source }) { BackColor = Color.Azure, Font = boldFont });
            listView.Items.Add(new ListViewItem(new[] { "4", "Amount", record.Amount }) { BackColor = Color.Azure, Font = boldFont });
            listView.Items.Add(new ListViewItem(new[] { "5", "InMemOfForenames", record.InMemOfForenames }) { BackColor = Color.Azure, Font = boldFont });
            listView.Items.Add(new ListViewItem(new[] { "6", "InMemoryOfSurname", record.InMemoryOfSurname }) { BackColor = Color.Azure, Font = boldFont });

            var fromCodeIsUnknown = record.FromCode.ToLower().Trim() == "unknown";


            if (!record.ManuallySelectedAsTributeFund)
            {
                listView.Items.Add(new ListViewItem(new[] { "7", "RequiredFromCode", "IMCO" }) { BackColor = Color.Azure, Font = boldFont });
                listView.Items.Add(new ListViewItem(new[] { "8", "WebId", record.WebId }) { BackColor = Color.Azure, Font = boldFont });

                if (fromCodeIsUnknown)
                {

                    listView.Items.Add(new ListViewItem(new[] { "9", "RelationshipToCode", record.ToCode }) { BackColor = Color.Azure, Font = boldFont });
                    listView.Items.Add(new ListViewItem(new[] { "10", "WebId", record.WebId }) { BackColor = Color.Azure, Font = boldFont });
                }
                else
                {
                    listView.Items.Add(new ListViewItem(new[] { "9", "RelationshipFromCode", record.FromCode }) { BackColor = Color.Azure, Font = boldFont });
                    listView.Items.Add(new ListViewItem(new[] { "10", "WebId", record.WebId }) { BackColor = Color.Azure, Font = boldFont });
                    listView.Items.Add(new ListViewItem(new[] { "11", "RelationshipToCode", record.ToCode }) { BackColor = Color.Azure, Font = boldFont });
                    listView.Items.Add(new ListViewItem(new[] { "12", "WebId", record.WebId }) { BackColor = Color.Azure, Font = boldFont });
                }
            }
            else
            {
                if(record.TributeNumberValid)
                {
                    listView.Items.Add(new ListViewItem(new[] { "7", "TributeFundNumber", record.TributeFundNumber }) { BackColor = Color.Azure, Font = boldFont });
                    listView.Items.Add(new ListViewItem(new[] { "8", "DonorName - Skip if TF found", $"*{record.Title} {record.DonorName}*" }) { BackColor = Color.Azure, Font = boldFont });
                }
                else
                {
                    listView.Items.Add(new ListViewItem(new[] { "7", "DonorName - Skip if TF found", $"*{record.Title} {record.DonorName}*" }) { BackColor = Color.Azure, Font = boldFont });
                }
            }

            listView.Items.Add(new ListViewItem(new[] { "", "OTHER", "" }) { BackColor = Color.LightGray });
            listView.Items.Add(new ListViewItem(new[] { "", "DonorName", $"{record.Title} {record.DonorName}" }));
            listView.Items.Add(new ListViewItem(new[] { "", "InMemoryOf", record.InMemoryOf }));
            listView.Items.Add(new ListViewItem(new[] { "", "InMemoryOfOther", record.InMemoryOfOther }));
            listView.Items.Add(new ListViewItem(new[] { "", "PaymentMethod", record.PaymentMethod }));
            listView.Items.Add(new ListViewItem(new[] { "", "InMemFrom", record.InMemFrom }));
            listView.Items.Add(new ListViewItem(new[] { "", "SelectedAsTributeFund", record.ManuallySelectedAsTributeFund ? "Y" : "N" }));
            listView.Items.Add(new ListViewItem(new[] { "", "TributeFundNumber", record.TributeFundNumber }));
            listView.Items.Add(new ListViewItem(new[] { "", "TributeFundNumberInvalid", record.TributeNumberValid ? "Y" : "N" }));


            if (fromCodeIsUnknown)
            {
                listView.Items.Add(new ListViewItem(new[] { "", "RelationshipFromCode", record.FromCode }));
            }
            else if(record.ManuallySelectedAsTributeFund)
            {
                listView.Items.Add(new ListViewItem(new[] { "", "RelationshipFromCode", record.FromCode }));
                listView.Items.Add(new ListViewItem(new[] { "", "RelationshipToCode", record.ToCode }) );
            }

        }

        private void LoadRecords(string filePath)
        {
            try
            {
                using (var reader = new StreamReader(filePath))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var allRecords = csv.GetRecords<DonationRecord>().ToList();

                    foreach (var record in allRecords)
                    {
                        record.SetGender(record.Title);

                        if (record.DonorGender == "Unknown")
                        {
                            using (var genderForm = new GenderSelectionForm(record.DonorName))
                            {
                                if (genderForm.ShowDialog() == DialogResult.OK)
                                {
                                    record.UpdateGender(genderForm.SelectedGender);
                                }
                                else
                                {
                                    record.UpdateGender("Unknown"); // Handle cases where dialog is closed without a selection
                                }
                            }
                        }

                        if(record.Source.ToLower().Trim().Contains("tf"))
                        {
                            if (record.TributeNumberValid)
                            {
                                Clipboard.SetText($"*{record.TributeFundNumber}*");
                                record.ManuallySelectedAsTributeFund = MessageBox.Show($"Is *{record.DonorName}* {record.TributeFundNumber} a Tribute Fund? (Yes/No). Go to Search Organisations => Funds. The Number has been copied to clipboard. (Press no and name will be copied to value) ", "Tribute Fund Check", MessageBoxButtons.YesNo) == DialogResult.Yes;
                            }
                            Clipboard.SetText($"*{record.DonorName}*");
                            record.ManuallySelectedAsTributeFund = MessageBox.Show($"Is *{record.DonorName}* {record.TributeFundNumber} a Tribute Fund? (Yes/No). Go to Search Organisations => Funds. The wildcardname has been copied to clipboard.", "Tribute Fund Check", MessageBoxButtons.YesNo) == DialogResult.Yes;

                        }
                    }

                    recordsPP = allRecords.Where(r => r.PaymentMethod == "PP").ToList();
                    recordsCC = allRecords.Where(r => r.PaymentMethod == "CC").ToList();
                }

                currentIndexPP = -1;
                currentIndexCC = -1;
                DisplayBatchData("PP"); // Initially display first PP record
                DisplayBatchData("CC"); // Initially display first CC record (though CC tab might not be visible immediately)
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load CSV: {ex.Message}");
            }
        }

        private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Call DisplayCurrentRecord based on the newly selected tab.
            if (tabControl.SelectedTab == tabPP)
            {
                DisplayCurrentRecord("PP");
            }
            else if (tabControl.SelectedTab == tabCC)
            {
                DisplayCurrentRecord("CC");
            }
        }

        private void ListView_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (!e.IsSelected) return;

            ListView activeListView = sender as ListView;
            int selectedIndex = e.ItemIndex;  // Get the index of the currently selected item

            // Check if the selected item has a valid index (assuming that an invalid index means an empty or placeholder row)
            if (string.IsNullOrEmpty(activeListView.Items[selectedIndex].SubItems[0].Text))
            {
                // Look upwards for the closest valid index
                int validIndex = GetPreviousValidIndex(activeListView, selectedIndex);
                if (validIndex != -1)
                {
                    currentIndex = validIndex;
                    activeListView.Items[validIndex].Selected = true;
                    activeListView.EnsureVisible(validIndex);
                }
            }
            else
            {
                currentIndex = selectedIndex;  // Set the current index to the selected item's index if it's valid
            }

            // Copy content to clipboard and update tooltip
            var item = activeListView.SelectedItems[0];
            Clipboard.SetText(item.SubItems[2].Text);  // Assume the value is in the third column
            ResetAndShowTooltip(activeListView, $"Cycled to: {item.SubItems[1].Text} - {item.SubItems[2].Text}");

            if(item.SubItems[0].Text == "7")
            {
                var tfManauallySelected = activeListView.FindItemWithText("SelectedAsTributeFund");
                var validTributeNumber = activeListView.FindItemWithText("TributeFundNumberInvalid");

                if (tfManauallySelected?.SubItems[2].Text == "Y")
                {
                    MessageBox.Show("Go to Search Organisations -> Funds.", "Tribute Fund Warning", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    if (validTributeNumber?.SubItems[2].Text == "Y")
                    {
                        MessageBox.Show("Try The Tribute Number First. Then try the wildcard search. If neither are right treat as normal.", "Tribute Fund Warning", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                }

                
            }
        }

        private int GetPreviousValidIndex(ListView listView, int startIndex)
        {
            // Traverse upwards from the current index to find a row with a valid index
            for (int i = startIndex - 1; i >= 0; i--)
            {
                if (!string.IsNullOrEmpty(listView.Items[i].SubItems[0].Text))
                {
                    return i;  // Return the first valid index found
                }
            }
            return -1;  // Return -1 if no valid index is found
        }

        private void ResetAndShowTooltip(ListView listView, string text)
        {
            Point tooltipPosition = GetTooltipPosition();
            tooltip.Hide(listView);
            tooltip.Show(text, listView, tooltipPosition, 3000); // Shows for 3000 milliseconds or adjust as needed
        }

        private Point GetTooltipPosition()
        {
            Rectangle workingArea = Screen.PrimaryScreen.WorkingArea;
            int x = workingArea.Right - 20; // Padding from the right edge
            int y = workingArea.Bottom - 20; // Padding from the bottom edge

            return new Point(x, y);
        }



    }

    public class GenderSelectionForm : Form
    {
        private Label nameLabel;
        public string SelectedGender { get; private set; }

        public GenderSelectionForm(string donorName)
        {
            InitializeComponent(donorName);
        }

        private void InitializeComponent(string donorName)
        {
            this.ClientSize = new Size(300, 180);
            this.Text = "Select Gender";

            nameLabel = new Label() { Text = $"Specify gender for {donorName}:", Location = new Point(10, 10), Width = 280, Height = 20 };

            Button maleButton = new Button() { Text = "Male", Location = new Point(110, 40), Width = 80 };
            Button femaleButton = new Button() { Text = "Female", Location = new Point(110, 80), Width = 80 };
            Button unknownButton = new Button() { Text = "Unknown", Location = new Point(110, 120), Width = 80 };

            maleButton.Click += (sender, e) => { SelectedGender = "Male"; this.DialogResult = DialogResult.OK; };
            femaleButton.Click += (sender, e) => { SelectedGender = "Female"; this.DialogResult = DialogResult.OK; };
            unknownButton.Click += (sender, e) => { SelectedGender = "Unknown"; this.DialogResult = DialogResult.OK; };

            this.Controls.Add(nameLabel);
            this.Controls.Add(maleButton);
            this.Controls.Add(femaleButton);
            this.Controls.Add(unknownButton);
        }
    }
}