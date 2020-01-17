//12/30/19 - v1 contains Save Files, Validation and Tree Builder
//12/30/19 - v1 complete, released to Beau for testing
//1/14/20 - v2.1 fixed single save file issue (zeroing out info in GO95 due to for loop enclosing too much of the code)

using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using PPL_Lib;
using SpreadsheetLight;
using MigraDoc.DocumentObjectModel;
using WeifenLuo.WinFormsUI.Docking;
using WeifenLuo.WinFormsUI.ThemeVS2015;
using PdfSharp;
using System.Xml;
using PPL_FiniteElementAnalysis;
using Microsoft.VisualBasic;
using System.Threading;
using System.Threading.Tasks;
using System.Text;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Collections.Specialized;
using System.Net;
using System.Net.Mail;
using System.Linq;
using System.Timers;



///----------------------------------------------------------------------------
/// Single Plugin created to consolidate all previous scripts
///----------------------------------------------------------------------------

namespace OCalcProPlugin
{
    public class Plugin : PPLPluginInterface
    {

        /// <summary>
        /// THis is the handle to the main O-Calc Pro component
        /// </summary>
        PPL_Lib.PPLMain cPPLMain = null;


        /// <summary>
        /// Declare the type of plugin as one of:
        ///         DOCKED_TAB
        ///         MENU_ITEM
        ///         BOTH_DOCKED_AND_MENU
        ///         CLEARANCE_SAG_PROVIDER
        /// </summary>
        public PLUGIN_TYPE Type
        {
            get
            {
                return PLUGIN_TYPE.MENU_ITEM;
            }
        }

        /// <summary>
        /// Declare the name of the plugin usd for synthesizing the registry keys ect
        /// </summary>
        public String Name
        {
            get
            {
                return "BCF_Master_Plugin";
            }
        }

        /// <summary>
        /// Optionally declare a description string (defaults to the name);
        /// </summary>
        public String Description
        {
            get
            {
                return Name;
            }
        }

        /// <summary>
        /// Add a tabbed form to the tabbed window (if the plugin type is 
        /// PLUGIN_TYPE.DOCKED_TAB 
        /// or
        /// PLUGIN_TYPE.BOTH_DOCKED_AND_MENU
        /// </summary>
        /// <param name="pPPLMain"></param>
        public void AddForm(PPL_Lib.PPLMain pPPLMain)
        {
            cPPLMain = pPPLMain;
            cForm = new PluginForm();
            Guid guid = new Guid(0x123eb510, 0xadcc, 0x4338, 0xa8, 0x12, 0x67, 0x6f, 0x32, 0xdb, 0x1e, 0x1e);



            cForm.cGuid = guid;
            cPPLMain.cDockedPanels.Add(cForm.cGuid.ToString(), cForm);
            foreach (Control ctrl in cPPLMain.Controls)
            {
                if (ctrl is WeifenLuo.WinFormsUI.Docking.DockPanel)
                {
                    cForm.Show(ctrl as WeifenLuo.WinFormsUI.Docking.DockPanel, WeifenLuo.WinFormsUI.Docking.DockState.Document);
                }
            }



            cForm.Show();
        }

        PluginForm cForm = null;

        class PluginForm : WeifenLuo.WinFormsUI.Docking.DockContent
        {
            public Guid cGuid;
            protected override string GetPersistString()
            {
                return cGuid.ToString();
            }
        }

        /// <summary>
        /// Perform clearance analysis if type is PLUGIN_TYPE.CLEARANCE_SAG_PROVIDER
        /// </summary>
        /// <param name="pMain"></param>
        /// <returns></returns>
        public PPLClearance.ClearanceSagProvider GetClearanceSagProvider(PPL_Lib.PPLMain pMain)
        {
            System.Diagnostics.Debug.Assert(Type == PLUGIN_TYPE.CLEARANCE_SAG_PROVIDER, Name + " is not a clearance provider plugin.");
            return null;
        }


        //Declare BCF dropdown and dropdown menu items
        ToolStripDropDownButton bcfButton = null; //main dropdown button for all plugins
        ToolStripMenuItem saveFilesBtn = null; //BCF Save Files button
        ToolStripMenuItem valBtn = null; //BCF Validation button
        ToolStripMenuItem treeBtn = null; //BCF Tree Attachment Pole Builder button
        ToolStripMenuItem coordBtn = null; //BCF Update Coordinates Button
        ToolStripMenuItem pgeDebugBtn = null;


        public void AddToMenu(PPL_Lib.PPLMain pPPLMain, System.Windows.Forms.ToolStrip pToolStrip)
        {
            //save the reference to the O-Calc Pro main
            cPPLMain = pPPLMain;


            //create the toolstrip buttons
            saveFilesBtn = new ToolStripMenuItem("BCF Save Files");
            saveFilesBtn.AutoToolTip = true;
            saveFilesBtn.ToolTipText = Description;
            saveFilesBtn.Click += SaveFilesBtn_Click;

            valBtn = new ToolStripMenuItem("BCF Validation");
            valBtn.AutoToolTip = true;
            valBtn.ToolTipText = Description;
            valBtn.Click += ValBtn_Click;

            treeBtn = new ToolStripMenuItem("BCF Tree Attachment Pole Builder");
            treeBtn.AutoToolTip = true;
            treeBtn.ToolTipText = Description;
            treeBtn.Click += TreeBtn_Click;

            coordBtn = new ToolStripMenuItem("BCF Update Coordinates Single");
            coordBtn.AutoToolTip = true;
            coordBtn.ToolTipText = "Updates the FAA coordinates for the current pole.  You must have the planning data spreadsheet saved to the desktop of your PGE computer";
            coordBtn.Click += CoordBtn_Click;

            pgeDebugBtn = new ToolStripMenuItem("PGE Debug");
            pgeDebugBtn.AutoToolTip = true;
            pgeDebugBtn.ToolTipText = "Debug";
            pgeDebugBtn.Click += PgeDebugBtn_Click;



            //create the dropdown button
            bcfButton = new ToolStripDropDownButton();
            bcfButton.Text = "BCF Plugins";
            bcfButton.DropDownDirection = ToolStripDropDownDirection.Default;
            bcfButton.DropDownOpened += BcfButton_DropDownOpened;
            bcfButton.Click += BcfButton_Click;

            //add toolstrip buttons to dropdown
            bcfButton.DropDownItems.Add(saveFilesBtn);
            bcfButton.DropDownItems.Add(valBtn);
            bcfButton.DropDownItems.Add(treeBtn);
            bcfButton.DropDownItems.Add(coordBtn);

            bcfButton.DropDownItems.Add(pgeDebugBtn);
            //code to enable some stuff for PGE
            int pgeEnabled = 69;
            pToolStrip.Items.Add(bcfButton);

        }

        private void PgeDebugBtn_Click(object sender, EventArgs e)
        {
            try
            {
                PgeDebug();
            }
            catch (Exception ex)
            {
                PPLMessageBox.Show(ex);
            }
        }

        private void CoordBtn_Click(object sender, EventArgs e)
        {
            try
            {
                UpdateCoordinatesSingle();
            }
            catch (Exception ex)
            {
                PPLMessageBox.Show(ex);
            }
        }


        private void TreeBtn_Click(object sender, EventArgs e)
        {
            try
            {
                TreeBuilder();
            }
            catch (Exception ex)
            {
                PPLMessageBox.Show(ex);
            }


        }

        private void ValBtn_Click(object sender, EventArgs e)
        {
            try
            {
                Validation();
            }
            catch (Exception ex)
            {
                PPLMessageBox.Show(ex);
            }

        }

        public void UpdateProgressBar(ProgressBar Bar, int Width)
        {
            
        }

        private void SaveFilesBtn_Click(object sender, EventArgs e)
        {
           
            try
            {
                SaveFiles();
            }
            catch (Exception ex)
            {
                PPLMessageBox.Show(ex);
            }
        }

        private void BcfButton_Click(object sender, EventArgs e)
        {
            bcfButton.DropDown.Width = 10000;
            bcfButton.DropDown.Show();
        }

        private void BcfButton_DropDownOpened(object sender, EventArgs e)
        {
            bool enabled = false;
            if (cPPLMain != null)
            {
                enabled = (cPPLMain.GetMainStructure() is PPLPole);
            }
            saveFilesBtn.Enabled = enabled;
            valBtn.Enabled = enabled;
            treeBtn.Enabled = enabled;
            coordBtn.Enabled = enabled;
        }




        public PPLElement FindElement(string pGuid, List<TreeNode> pNodes)
        {
            if (pNodes == null)
            {
                return (PPLElement)null;
            }
            foreach (TreeNode pNode in pNodes)
            {
                if (pNode.Tag is PPLElement)
                {
                    PPLElement tag = (PPLElement)pNode.Tag;
                    string elemGuid = tag.Guid.ToString();
                    if (elemGuid == pGuid)
                    {
                        PPLElement pplElement = tag.CopyElement();
                        pplElement.Parent = (PPLElement)null;
                        pplElement.Children.Clear();
                        return pplElement;
                    }
                }

                if (pNode.Nodes != null && pNode.Nodes.Count > 0)
                {
                    PPLElement defaultElement = FindElementPartDeux(pGuid, pNode.Nodes);
                    if (defaultElement != null)
                    {
                        return defaultElement;
                    }

                }
            }
            return (PPLElement)null;
        }

        //argument passed from FindElement is a TreeNodeCollection
        public PPLElement FindElementPartDeux(string pGuid, TreeNodeCollection pNodes)
        {
            if (pNodes == null)
            {
                return (PPLElement)null;
            }
            foreach (TreeNode pNode in pNodes)
            {
                if (pNode.Tag is PPLElement)
                {
                    PPLElement tag = (PPLElement)pNode.Tag;
                    string elemGuid = tag.Guid.ToString();
                    if (elemGuid == pGuid)
                    {
                        PPLElement pplElement = tag.CopyElement();
                        pplElement.Parent = (PPLElement)null;
                        pplElement.Children.Clear();
                        return pplElement;
                    }
                }

                if (pNode.Nodes != null && pNode.Nodes.Count > 0)
                {
                    PPLElement defaultElement = FindElementPartDeux(pGuid, pNode.Nodes);
                    if (defaultElement != null)
                    {
                        return defaultElement;
                    }

                }
            }
            return (PPLElement)null;
        }

        public PPLElement FindCase(string pName, List<TreeNode> pNodes)
        {
            if (pNodes == null)
            {
                return (PPLElement)null;
            }
            foreach (TreeNode pNode in pNodes)
            {
                if (pNode.Tag is PPLElement)
                {
                    PPLElement tag = (PPLElement)pNode.Tag;
                    string elemName = tag.GetValueAndConvertToString("Name");
                    if (elemName.ToUpper() == pName.ToUpper())
                    {
                        PPLElement pplElement = tag.CopyElement();
                        pplElement.Parent = (PPLElement)null;
                        pplElement.Children.Clear();
                        return pplElement;
                    }
                }

                if (pNode.Nodes != null && pNode.Nodes.Count > 0)
                {
                    PPLElement defaultElement = FindCasePartDeux(pName, pNode.Nodes);
                    if (defaultElement != null)
                    {
                        return defaultElement;
                    }

                }
            }
            return (PPLElement)null;
        }

        //argument passed from FindElement is a TreeNodeCollection
        public PPLElement FindCasePartDeux(string pName, TreeNodeCollection pNodes)
        {
            if (pNodes == null)
            {
                return (PPLElement)null;
            }
            foreach (TreeNode pNode in pNodes)
            {
                if (pNode.Tag is PPLElement)
                {
                    PPLElement tag = (PPLElement)pNode.Tag;
                    string elemName = tag.GetValueAndConvertToString("Name");
                    if (elemName.ToUpper() == pName.ToUpper())
                    {
                        PPLElement pplElement = tag.CopyElement();
                        pplElement.Parent = (PPLElement)null;
                        pplElement.Children.Clear();
                        return pplElement;
                    }
                }

                if (pNode.Nodes != null && pNode.Nodes.Count > 0)
                {
                    PPLElement defaultElement = FindCasePartDeux(pName, pNode.Nodes);
                    if (defaultElement != null)
                    {
                        return defaultElement;
                    }

                }
            }
            return (PPLElement)null;
        }

        public void OcalcReload(PPLMain pPPLmain)
        {
            cPPLMain = pPPLmain;
            cPPLMain.ForceReloadDEP();
            cPPLMain.Rebuild3DDisplay();
            cPPLMain.UpdateCapacitySummary();
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        public string GetActiveWindowsTXT()
        {
            try
            {
                int chars = 256;
                StringBuilder buff = new StringBuilder(chars);
                IntPtr handle = GetForegroundWindow();
                string RET = "";
                if (handle != null)
                {
                    if (GetWindowText(handle, buff, chars) > 0)
                    {
                        RET = buff.ToString();
                        //listBox2.Items.Add(handle.ToString());                   
                    }
                }

                return RET;
            }
            catch
            {
                return "NO Windows Active";
            }
        }


        public void SavePPLX(PPLMain pPPLMain, string pFileName)
        {
            pPPLMain.DoSavePole(pFileName, pPPLMain.SaveAsVersion, true, true);

        }

        public void SaveCustomReport(PPLMain pPPLMain, string pOsmoseCustomReportFileName, string pOutputPath)
        {
            PPLExcelReportsForm excelReport = new PPLExcelReportsForm(pPPLMain);
            excelReport.CreateExcelReport(pOsmoseCustomReportFileName, pOutputPath);
        }

        public void SaveGO95(PPLMain pPPLMain, PPLPole pPole, string pFileName)
        {
            PPLReportsForm reportTemplate = new PPLReportsForm(pPPLMain, pPole);
            reportTemplate.BuildReportsList();  //don't worry about the dock content error. it still works in ocalc
            MigraDocReportControl mgDocControl = new MigraDocReportControl();
            foreach (ReportTemplate rep in reportTemplate.cAvailableReports)
            {
                if (Convert.ToString(rep.Description).Equals("GO95 Analysis Report"))
                {
                    rep.CreateReport(pPPLMain, null);
                    rep.AddPole(pPole, (string)null);

                    //mgDocControl.SetReport(doc, pPPLMain, (PPLWorkingDialog)null);

                    Document doc = rep.GetReport();
                    MigraDoc.Rendering.PdfDocumentRenderer pdfRenderer = new MigraDoc.Rendering.PdfDocumentRenderer(false, PdfSharp.Pdf.PdfFontEmbedding.Always);
                    pdfRenderer.Document = doc;
                    pdfRenderer.RenderDocument();
                    pdfRenderer.PdfDocument.Save(pFileName);
                    rep.Reset();
                    //mgDocControl.SetReport(doc, pPPLMain, (PPLWorkingDialog)null);
                }
            }
        }



        /// //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public void Validation()
        {
            //Form currForm = Form.ActiveForm;

            //Control.ControlCollection coll = currForm.Controls;
            //PPLMain pMain = currForm.ActiveControl as PPLMain;
            //foreach (Control ctl in coll)
            //{

            //}

            //pMain.Focus();
            //FormCollection frmColl = Application.OpenForms;
            //foreach (Form f in frmColl)
            //{
            //    if (f.Name.Contains("PGE"))
            //    {
            //        PPLMain main = f.ActiveControl as PPLMain;
            //    }
            //}

            //cPPLMain = pPPLMain;
            PPLPole pole = cPPLMain.GetPole();
            cPPLMain.cInventory.SortInventory();
            string sapPM = pole.GetValueString("Aux Data 1");
            string ocalcPoleLoc = pole.GetValueString("Aux Data 3");
            double poleSetDepth = pole.GetValueDouble("BuryDepthInInches");
            double upperCommHoa = 0;
            string owner = null;
            string[] allowableCompositePoles = new string[] { "T1D45", "T3D45", "H2D45", "H4D45", "T1D50", "T3D50", "H2D50", "H4D50", "T1D55", "T3D55", "H2D55", "H4D55", "H3D65", "H5D65" };      //these are the poles we can order in EES
            double lowestCrossArmHeight = pole.GetValueDouble("LengthInInches");  //initialize to height of pole.  check each iteration of crossarm to find the lowest
            double[] poleSettingDepth = new double[] { (4.5 * 12), (5 * 12), (5.5 * 12), (6 * 12), (6.5 * 12), (7 * 12), (7.5 * 12), (8 * 12), (8.5 * 12), (9 * 12), (9.5 * 12), (10 * 12), (10.5 * 12) };
            double[] allowableMoment = new double[] { 8050, 11000, 17300, 24500, 35300, 47500, 67000, 91000, 126000, 170000, 225000, 295000, 395000 };
            PPLElementCollection elems = pole.GetElementList();

            //check for undefined owners & get highest comm height of attachment & check for composite poles that can't be ordered & 
            //update composite pole overturn moment
            foreach (PPLElement elem in elems)
            {

                //check for element with undefined owner
                owner = elem.GetValueString("Owner");
                if (owner == "<Undefined>")
                {
                    MessageBox.Show("The owner of " + elem.DescriptionString + " is Undefined.");
                }

                //set undefined owner of anchor to PG&E
                if (owner == "<Undefined>" && elem is PPLAnchor)
                {
                    elem.SetValue("Owner", "PG&E");
                }

                //check for composite pole that can't be ordered
                if (elem is PPLCompositePole)
                {
                    string poleCatalogName = elem.GetValueString("CatalogName");
                    bool isAllowed = false;
                    for (int i = 0; i < allowableCompositePoles.Length; i++)
                    {
                        isAllowed = allowableCompositePoles[i].Equals(poleCatalogName);
                        if (isAllowed)
                        {
                            break;
                        }
                        else if (i == allowableCompositePoles.Length - 1)
                        {
                            PPLMessageBox.Show("Pole cannot be ordered in EES. Please choose a different pole.");
                        }
                    }

                    //update allowable moment if pole setting depth is manually changed
                    for (int j = 0; j < allowableMoment.Length; j++)
                    {
                        double currentPoleSetDepth = elem.GetValueDouble("BuryDepthInInches");
                        double currentOverturnMoment = elem.GetValueDouble("OverturnMoment");
                        if (Array.IndexOf(poleSettingDepth, currentPoleSetDepth).Equals(Array.IndexOf(allowableMoment, currentOverturnMoment)))
                        {
                            break;
                        }
                        else if (currentPoleSetDepth == poleSettingDepth[j])
                        {
                            elem.SetValue("OverturnMoment", allowableMoment[j]);
                            PPLMessageBox.Show("Overturn moment has been updated to match current set depth");
                            break;
                        }
                        else if (j == allowableMoment.Length - 1)
                        {
                            PPLMessageBox.Show("Could not find an overturn moment to match " + currentPoleSetDepth / 12 + "' set depth");
                            break;
                        }
                    }
                }

                //check for comm out of grade
                if (elem is PPLSpan && elem.Parent is PPLInsulator && (owner == "Comm" | owner == "COMM" | owner == "comm"
                                                                       | owner == "CATV" | owner == "catv" | owner == "Catv"))
                {
                    double commHoa = (elem.Parent.GetValueDouble("CoordinateZ") - poleSetDepth) / 12;
                    if (commHoa > upperCommHoa)
                    {
                        upperCommHoa = commHoa;
                    }
                }

                //check for wood crossarm
                if (elem is PPLCrossArm)
                {
                    string crossarmMaterial = elem.GetValueString("Material");
                    if (crossarmMaterial == "Wood" | crossarmMaterial == "Other")
                    {
                        MessageBox.Show("The " + elem.DescriptionString + " crossarm material is not currently allowed.");
                    }

                    //get height of lowest crossarm for transformer mounting height reference
                    double currentCrossarmHeight = elem.GetValueDouble("CoordinateZ");
                    if (currentCrossarmHeight <= lowestCrossArmHeight)
                    {
                        lowestCrossArmHeight = currentCrossarmHeight;
                    }
                    //PPLMessageBox.Show(Convert.ToString(lowestCrossArmHeight));
                }

                //check that transformer is mounted 3.25' below lowest primary crossarm and set mounting height if it is not
                if (/*elem is PPLGenericEquipment || */elem is PPLTransformer && elem.TypeString == "Transformer")
                {
                    double txHeight = elem.GetValueDouble("CoordinateZ");
                    if (txHeight != (lowestCrossArmHeight - (3.25 * 12)))
                    {
                        DialogResult dialogResult = MessageBox.Show("Current height of transformer is not 3.25' below lower crossarm.  Do you want to set it?", "Transformer Height Notification", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                        {
                            elem.SetValue("CoordinateZ", lowestCrossArmHeight - (3.25 * 12));
                            MessageBox.Show("Transformer height has been updated");
                        }
                        else
                        {
                            MessageBox.Show("Transformer height has not been updated");
                        }
                    }
                }
            }


            //MessageBox.Show("the highest comm hoa is " + upperCommHoa.ToString());

            //check for pge owner assigned guys below comm owned spans
            foreach (PPLElement guy in elems)
            {
                if (guy is PPLGuyBrace)
                {
                    double guyHoa = (guy.GetValueDouble("CoordinateZ") - poleSetDepth) / 12;
                    string guyOwner = guy.GetValueString("Owner");
                    if (guyHoa <= upperCommHoa && guyOwner != "Comm")
                    {
                        MessageBox.Show("The " + guy.DescriptionString + " guy attached at " + guyHoa + "ft, is in communication's space, but the owner is " + guyOwner + ".");
                        guy.SetValue("Owner", "Comm");
                    }

                    //Check for comm owned guys above comm owned spans
                    else if (guyHoa > upperCommHoa && guyOwner == "Comm")
                    {
                        PPLMessageBox.Show("The " + guy.DescriptionString + " guy attached at " + guyHoa + "ft is above communication's space, but the owner is " + guyOwner + ".");
                        guy.SetValue("Owner", "PG&E");
                    }
                }
            }
            OcalcReload(cPPLMain);
        }


        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public void SaveFiles()
        {
            Form currForm = Form.ActiveForm;

            Form1 frm = new Form1();
            frm.Text = "Saving Files";
            frm.Show();
            foreach (Control ct in frm.Controls)
            {
                if (ct is ProgressBar)
                {
                    ProgressBar bar = ct as ProgressBar;
                    bar.Increment(50);
                }
            }
            //pbar.Width = 700;
            //int width = pbar.Width;
            //frm.Controls.Add(pbar);
            //frm.Show();


            PPLPole pole = cPPLMain.GetPole();
            List<TreeNode> treeNodeList = cPPLMain.cCatalogManager.GetNodes(PPLCatalogForm.CATALOG_TYPE.MASTER);
            List<PPLEnvironment> envList = pole.GetEnvironments();
            string sapPM = pole.GetValueString("Aux Data 1");
            string ocalcPoleLoc = pole.GetValueString("Aux Data 3");


            //get existing HFTD load case and replace it from the treelist to fix NESC issue
            foreach (PPLEnvironment env in envList)
            {
                string loadCaseName = env.GetValueAndConvertToString("Name");
                string loadCaseCode = env.GetValueAndConvertToString("Method");
                if (loadCaseName.Contains("HFTD") && loadCaseCode == "NESC")
                {
                    PPLEnvironment replaceLoadCase = FindCase(loadCaseName, treeNodeList) as PPLEnvironment;
                    //in the PPL_Lib library the Substitute method takes two arguments, PPLElement and PPLMain
                    //but this gives an error in the script.  When using this script for testing you have to comment out
                    //the pPPLMain.  But when using it on the PGE computer you must uncomment it or it will throw an error
                    env.Substitute(replaceLoadCase/*, cPPLMain*/);
                    //pPPLMain.SelectLoadCase(env);
                    OcalcReload(cPPLMain);
                }
            }
            //define folder names
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string desktopPMFolder = desktopPath + @"\" + sapPM + " Ocalc Files";
            string pplxFolder = desktopPMFolder + "\\PPLX";
            string pplxFileName = pplxFolder + "\\" + ocalcPoleLoc + ".pplx";
            string customReportFolder = desktopPMFolder + "\\Custom Reports";
            string customReportFileName = customReportFolder + "\\" + ocalcPoleLoc + ".xlsx";
            string go95Folder = desktopPMFolder + "\\GO95 Analysis";
            string go95FileName = go95Folder + "\\" + ocalcPoleLoc + ".pdf";
            string osmosePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData); //location of custom report excel file
            string osmoseCustomReportFileName = osmosePath + "\\OsmosePPL\\Reports\\BCF Report v1.4.xlsx";

            //create folders from folder names
            System.IO.Directory.CreateDirectory(desktopPMFolder);
            System.IO.Directory.CreateDirectory(pplxFolder);
            System.IO.Directory.CreateDirectory(customReportFolder);
            System.IO.Directory.CreateDirectory(go95Folder);

            SavePPLX(cPPLMain, pplxFileName);
            
            SaveGO95(cPPLMain, pole, go95FileName);

            SaveCustomReport(cPPLMain, osmoseCustomReportFileName, customReportFileName);







            PPLMessageBox.Show("Files Saved!");
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public void TreeBuilder()
        {
            /////////////Catalog Guids and other constants//////////////////
            const string spoolGuid = "21670682-ab38-484c-bed8-a212650ee25a";
            const string dummyLoadGuid = "85a21fd2-86ad-4744-b1eb-fc6fb232596b";
            const double pi = Math.PI;

            /////////////End Catalog Guids and other constants//////////////

            PPLPole pole = cPPLMain.GetPole();
            List<TreeNode> treeNodeList = cPPLMain.cCatalogManager.GetNodes(PPLCatalogForm.CATALOG_TYPE.MASTER);
            //dummy load case needs to be added to the pole in order to populate the Aux Data fields
            PPLEnvironment dummyLoadCase = FindElement(dummyLoadGuid, treeNodeList) as PPLEnvironment;
            pole.AddChild(dummyLoadCase);
            string sapPM = pole.GetValueString("Aux Data 1");
            string poleLoc = pole.GetValueString("Aux Data 3");

            //Get pole info from planning spreadsheet
            string pdFilePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + sapPM + " Planning Data.xlsx";
            string treeAttachSheet = "Ocalc Tree Attachments";

            SLDocument planningData = new SLDocument(pdFilePath, treeAttachSheet);
            SLWorksheetStatistics wsStats = planningData.GetWorksheetStatistics();
            int numRows = wsStats.NumberOfRows;
            int numCols = wsStats.NumberOfColumns;
            string instructionStr = string.Empty;
            for (int i = 2; i <= numRows; i++)
            {
                if (planningData.GetCellValueAsString(i, 1).Contains(poleLoc))
                //if (planningData.GetCellValueAsString(i, 1).Equals(poleLoc))
                {
                    instructionStr = planningData.GetCellValueAsString(i, 2);
                }
            }
            planningData.CloseWithoutSaving();

            //assign variables from instruction array
            string[] instructionArr = instructionStr.Split(',');
            string hftdGuid = instructionArr[0];

            string go95Guid = instructionArr[1];
            string wireGuid = instructionArr[2];
            double[] distanceArr = instructionArr[3].Split(':').Select(Double.Parse).ToArray();

            //check for spans over 125' and exit program if found
            for (int i = 0; i <= distanceArr.Length - 1; i++)
            {
                if (distanceArr[i] > 125)
                {
                    PPLMessageBox.Show("One or more spans to this pole exceeds 125'. Interset is required", "Cannot Build Pole!");
                    goto end;
                }
            }


            double[] headingArr = instructionArr[4].Split(':').Select(Double.Parse).ToArray();


            ////if length of distanceArr != 2 ... need to figure out logic for this though the case would be very rare
            //if (distanceArr.Length == 2)
            //{
            //    double buryDepth = 7 * 12;
            //    pole.SetValue("BuryDepthInInches", buryDepth);
            //    PPLEnvironment hftdLoadCase = FindElement(hftdGuid, treeNodeList) as PPLEnvironment;
            //    PPLEnvironment go95LoadCase = FindElement(go95Guid, treeNodeList) as PPLEnvironment;

            //    //in the PPL_Lib library the Substitute method takes two arguments, PPLElement and PPLMain
            //    //but this gives an error in the script.  When using this script for testing you have to comment out
            //    //the pPPLMain.  But when using it on the PGE computer you must uncomment it or it will throw an error
            //    dummyLoadCase.Substitute(hftdLoadCase, pPPLMain);
            //    //pole.AddChild(hftdLoadCase);
            //    pole.AddChild(go95LoadCase);

            //    PPLInsulator spool = FindElement(spoolGuid, treeNodeList) as PPLInsulator;
            //    double deltaAngle = Math.Abs(headingArr[1] - headingArr[0]);
            //    if (deltaAngle > 180)
            //    {
            //        spool.CoordinateA = ((Math.Max(headingArr[0], headingArr[1]) - Math.Min(headingArr[0], headingArr[1])) / 2 + Math.Min(headingArr[0], headingArr[1]) + 180) * pi / 180;
            //    }

            //    else
            //    {
            //        spool.CoordinateA = ((Math.Max(headingArr[0], headingArr[1]) - Math.Min(headingArr[0], headingArr[1])) / 2 + Math.Min(headingArr[0], headingArr[1])) * (pi / 180);
            //    }
            //    spool.CoordinateZ = pole.LengthInInches - 9;
            //    spool.SetValue("Side", "Inline");
            //    pole.AddChild(spool);
            //    spool.SnapToParent();

            //    PPLSpan[] spans = new PPLSpan[distanceArr.Length];
            //    for (int i = 0; i <= distanceArr.Length - 1; i++)
            //    {
            //        spans[i] = FindElement(wireGuid, treeNodeList) as PPLSpan;
            //        spans[i].CoordinateA = (headingArr[i] * pi / 180) - spool.CoordinateA;
            //        spans[i].SetValue("SpanDistanceInInches", distanceArr[i] * 12);
            //        spool.AddChild(spans[i]);
            //        spans[i].SnapToInsulator();
            //    }
            //}
            //else
            //{
            //    PPLMessageBox.Show("Cannot build pole!");
            //}

            double buryDepth = 6.5 * 12;
            pole.SetValue("BuryDepthInInches", buryDepth);
            PPLEnvironment hftdLoadCase = FindElement(hftdGuid, treeNodeList) as PPLEnvironment;
            PPLEnvironment go95LoadCase = FindElement(go95Guid, treeNodeList) as PPLEnvironment;

            //in the PPL_Lib library the Substitute method takes two arguments, PPLElement and PPLMain
            //but this gives an error in the script.  When using this script for testing you have to comment out
            //the pPPLMain.  But when using it on the PGE computer you must uncomment it or it will throw an error
            dummyLoadCase.Substitute(hftdLoadCase/*, cPPLMain*/);
            //pole.AddChild(hftdLoadCase);
            pole.AddChild(go95LoadCase);


            double minHeading;
            double maxHeading;
            if (distanceArr.Length < 4)
            {
                PPLInsulator spool = FindElement(spoolGuid, treeNodeList) as PPLInsulator;

                if (distanceArr.Length == 2)
                {
                    minHeading = Math.Min(headingArr[0], headingArr[1]);
                    maxHeading = Math.Max(headingArr[0], headingArr[1]);
                }

                else if (distanceArr.Length == 3)
                {
                    minHeading = headingArr[0];
                    maxHeading = headingArr[0];
                    for (int i = 0; i < headingArr.Length - 1; i++)
                    {
                        if (minHeading > headingArr[i])
                        {
                            minHeading = headingArr[i];
                        }
                        else if (maxHeading < headingArr[i])
                        {
                            maxHeading = headingArr[i];
                        }
                    }
                }
                else
                {
                    PPLMessageBox.Show("Cannot build pole!");
                    goto end;
                }

                double deltaAngle = maxHeading - minHeading;
                if (deltaAngle > 180)
                {
                    spool.CoordinateA = ((deltaAngle / 2) + minHeading + 180) * pi / 180;
                }
                else
                {
                    spool.CoordinateA = ((deltaAngle / 2) + minHeading) * (pi / 180);
                }

                spool.CoordinateZ = pole.LengthInInches - 9;
                spool.SetValue("Side", "Inline");
                pole.AddChild(spool);
                spool.SnapToParent();

                PPLSpan[] spans = new PPLSpan[distanceArr.Length];
                for (int i = 0; i <= distanceArr.Length - 1; i++)
                {
                    spans[i] = FindElement(wireGuid, treeNodeList) as PPLSpan;
                    spans[i].CoordinateA = (headingArr[i] * pi / 180) - spool.CoordinateA;
                    spans[i].SetValue("SpanDistanceInInches", distanceArr[i] * 12);
                    spool.AddChild(spans[i]);
                    spans[i].SnapToInsulator();
                }
            }

            else if (distanceArr.Length == 4)
            {


                PPLInsulator[] spools = new PPLInsulator[2];
                PPLSpan[] multiSpans = new PPLSpan[headingArr.Length];

                double[] minHead = new double[2];
                minHead[0] = Math.Min(headingArr[0], headingArr[2]);
                minHead[1] = Math.Min(headingArr[1], headingArr[3]);

                double[] maxHead = new double[2];
                maxHead[0] = Math.Max(headingArr[0], headingArr[2]);
                maxHead[1] = Math.Max(headingArr[1], headingArr[3]);

                double[] deltaAng = new double[2];
                //deltaAng[0] = Math.Abs(headingArr[0] - headingArr[2]);
                //deltaAng[1] = Math.Abs(headingArr[1] - headingArr[3]);

                int spansIndex = 0;
                int lessThan = 2;

                for (int i = 0; i <= spools.Length - 1; i++)
                {

                    deltaAng[i] = maxHead[i] - minHead[i];
                    spools[i] = FindElement(spoolGuid, treeNodeList) as PPLInsulator;
                    if (deltaAng[i] > 180)
                    {
                        spools[i].CoordinateA = ((deltaAng[i] / 2) + minHead[i] + 180) * pi / 180;
                    }
                    else
                    {
                        spools[i].CoordinateA = ((deltaAng[i] / 2) + minHead[i]) * (pi / 180);
                    }


                    spools[i].CoordinateZ = pole.LengthInInches - 9;
                    spools[i].SetValue("Side", "Inline");
                    pole.AddChild(spools[i]);
                    spools[i].SnapToParent();
                    for (int j = spansIndex; j <= lessThan; j += 2)
                    {
                        multiSpans[j] = FindElement(wireGuid, treeNodeList) as PPLSpan;
                        multiSpans[j].CoordinateA = (headingArr[j] * pi / 180) - spools[i].CoordinateA;
                        multiSpans[j].SetValue("SpanDistanceInInches", distanceArr[j] * 12);
                        spools[i].AddChild(multiSpans[j]);
                        multiSpans[j].SnapToInsulator();
                    }
                    spansIndex++;
                    lessThan++;

                }



            }
            OcalcReload(cPPLMain);
        end:
            { }

        }

        public ToolStripMenuItem GetSaveAsProposedDesignButton()
        {
            ToolStripMenuItem saveAsProposed = null;
            foreach (Control ctl in cPPLMain.Controls)
            {
                if (ctl.Name == "toolStrip1")
                {
                    ToolStrip strip = ctl as ToolStrip;
                    foreach (var button in strip.Items)
                    {
                        if (button is ToolStripDropDownButton)
                        {
                            ToolStripDropDownButton dButton = button as ToolStripDropDownButton;
                            if (dButton.Name == "toolStripDropDownButton1")
                            {
                                ToolStripDropDownButton editButton = dButton as ToolStripDropDownButton;
                                foreach (var it in editButton.DropDownItems)
                                {
                                    if (it is ToolStripMenuItem)
                                    {
                                        ToolStripMenuItem menuItem = it as ToolStripMenuItem;
                                        if (menuItem.Text == "Save PLC As")
                                        {
                                            ToolStripMenuItem saveAs = menuItem as ToolStripMenuItem;
                                            foreach (var sv in saveAs.DropDownItems)
                                            {
                                                if (sv is ToolStripMenuItem)
                                                {
                                                    ToolStripMenuItem sv1 = sv as ToolStripMenuItem;
                                                    if (sv1.Text == "Proposed Design")
                                                    {
                                                        saveAsProposed = sv1;
                                                        return saveAsProposed;
                                                    }
                                                }
                                            }

                                        }

                                    }
                                }

                            }

                        }

                    }

                }
            }
            return null;
        }


        public ToolStripMenuItem GetFAAButton()
        {
            ToolStripMenuItem overrideFAA = null;
            foreach (Control ctl in cPPLMain.Controls)
            {
                if (ctl.Name == "toolStrip1")
                {
                    ToolStrip strip = ctl as ToolStrip;
                    foreach (var button in strip.Items)
                    {
                        if (button is ToolStripDropDownButton)
                        {
                            ToolStripDropDownButton dButton = button as ToolStripDropDownButton;
                            if (dButton.Name == "toolStripDropDownButton_Edit")
                            {
                                ToolStripDropDownButton editButton = dButton as ToolStripDropDownButton;
                                foreach (var it in editButton.DropDownItems)
                                {
                                    if (it is ToolStripMenuItem)
                                    {
                                        ToolStripMenuItem menuItem = it as ToolStripMenuItem;
                                        if (menuItem.Text == "Override Coordinates for &FAA")
                                        {
                                            overrideFAA = menuItem as ToolStripMenuItem;
                                            overrideFAA.Enabled = true;
                                            return overrideFAA;
                                        }

                                    }
                                }
                            }
                        }
                    }
                }
            }
            return null;
        }


        public void UpdateCoordinatesSingle()
        {
            //string current = GetActiveWindowsTXT();
            //Form currForm = Form.ActiveForm;

            //Control.ControlCollection coll = currForm.Controls;
            //PPLMain pMain = currForm.ActiveControl as PPLMain;
            //foreach (Control ctl in coll)
            //{

            //}

            //pMain.Focus();
            //FormCollection frmColl = Application.OpenForms;
            foreach (Form f in Application.OpenForms)
            {
                //if (f.Name.Contains("PGE"))
                //{
                //    PPLMain main = f.ActiveControl as PPLMain;
                //}
            }
            //PPLMain main = WeifenLuo.WinFormsUI.ThemeVS2013.VS2013DockPane as PPLMain;



            bool enabled = false;



            ToolStripMenuItem proposed = GetSaveAsProposedDesignButton();
            ToolStripMenuItem overrideFAA = GetFAAButton();
            enabled = overrideFAA.Enabled;


            if (enabled)
            {
                PPLPole pole = cPPLMain.GetPole();
                string sapPM = pole.GetValueString("Aux Data 1");
                string poleLoc = pole.GetValueString("Aux Data 3");
                string slLoc = poleLoc.Substring(poleLoc.LastIndexOf('_') + 1);
                string poleLat = string.Empty;
                string poleLong = string.Empty;

                string pdFilePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + sapPM + " Planning Data For PGE Computer.xlsx";
                string treeAttachSheet = "Existing Pole Data";


                SLDocument planningData = new SLDocument(pdFilePath, treeAttachSheet);
                SLWorksheetStatistics wsStats = planningData.GetWorksheetStatistics();
                int numRows = wsStats.NumberOfRows;
                int numCols = wsStats.NumberOfColumns;
                bool found = false;                for (int i = 2; i <= numRows; i++)
                {
                    if (planningData.GetCellValueAsString(i, 2).Equals(slLoc))
                    {
                        poleLat = planningData.GetCellValueAsString(i, 17);
                        poleLong = planningData.GetCellValueAsString(i, 18);
                        found = true;
                        PPLMessageBox.Show("Pole found");
                    }
                }

                if (!found)
                {
                    PPLMessageBox.Show("Pole not found");
                }
                planningData.CloseWithoutSaving();

            }
            else
            {
                PPLMessageBox.Show("Override Coordinates for FAA not enabled for this pole.");
            }

        }

        public void SideJob()
        {
            Form1 frm1 = new Form1();
            frm1.ShowDialog();


        }


        public void SaveAsProposedDesign()
        {
            //Send keys doesnt work with performClick
            //ToolStripMenuItem proposed = GetSaveAsProposedDesignButton();
            //proposed.PerformClick();
            SendKeys.Send("%f");
            SendKeys.Send("{DOWN}");
            SendKeys.Send("{RIGHT}");
            SendKeys.Send("{DOWN 2}");
            SendKeys.Send("{ENTER}");
            SendKeys.Send("{TAB}");
            SendKeys.Send("{ENTER}");
            //SendKeys.Send("{ENTER}");

        }

        public void OverrideFaa(/*string poleLat, string poleLong*/)
        {
            
            //The sendkeys commands dont seem to work after using perform click
            ToolStripMenuItem overrideFaaBtn = GetFAAButton();
            overrideFaaBtn.PerformClick();
            //Delay(5000);
            //SendKeys.Send("{ESC}");
            //SendKeys.Send("%e");
            //SendKeys.Send("{DOWN 7}");
            //SendKeys.Send("{ENTER}");
            //SendKeys.Send("{TAB 2}");
            //SendKeys.Send("^a");
            //SendKeys.Send("{BACKSPACE}");
            //SendKeys.Send(poleLat);
            //SendKeys.Send("{TAB}");
            //SendKeys.Send("^a");
            //SendKeys.Send("{BACKSPACE}");
            //SendKeys.Send(poleLong);
            //SendKeys.Send("{TAB 2}");
            //SendKeys.Send("{ENTER}");
            //SendKeys.Send("{ENTER}");
        }




        public void PgeDebug()
        {
            Thread th = new Thread(OverrideFaa);
            th.Start();

            ///////////////////
            ///reference to pplmain
            Thread.Sleep(30000);
            PPLMain pMain = null;
            foreach(Form f in Application.OpenForms)
            {
                //if (f.Name is "PGE_PPL")
                //{
                //    pMain = f.ActiveControl as PPLMain;
                //}

                //if (f.Name is "PGE_DataBlockForm")
                //{
                //    var pge = f;
                //    foreach(Control ctl in pge.Controls)
                //    {

                //    }
                    
                //}


            //    if  (f.Name is "Form1")
            //    {
            //        foreach (Control ctl in f.Controls)
            //        {

            //        }
            //    }
            }
            th.Abort();


        }

    }
}
