using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
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
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Collections.Specialized;
using System.Net;
using System.Net.Mail;


namespace OCalcProPlugin
{
    public partial class Form1 : Form
    {
        //PPL_Lib.PPLMain cPPLMain = null;


        /// <summary>
        /// Add a tabbed form to the tabbed window (if the plugin type is 
        /// PLUGIN_TYPE.DOCKED_TAB 
        /// or
        /// PLUGIN_TYPE.BOTH_DOCKED_AND_MENU
        /// </summary>
        /// <param name="pPPLMain"></param>

        //public void AddForm(PPL_Lib.PPLMain pPPLMain)
        //{
        //    cPPLMain = pPPLMain;
        //    cForm = new PluginForm();
        //    Guid guid = new Guid(0x123eb510, 0xadcc, 0x4338, 0xa8, 0x12, 0x67, 0x6f, 0x32, 0xdb, 0x1e, 0x1e);


        //    cForm.cGuid = guid;
        //    cPPLMain.cDockedPanels.Add(cForm.cGuid.ToString(), cForm);
        //    foreach (Control ctrl in cPPLMain.Controls)
        //    {
        //        if (ctrl is WeifenLuo.WinFormsUI.Docking.DockPanel)
        //        {
        //            cForm.Show(ctrl as WeifenLuo.WinFormsUI.Docking.DockPanel, WeifenLuo.WinFormsUI.Docking.DockState.Document);
        //        }
        //    }
        //    cForm.Show();
        //}

       // PluginForm cForm = null;

        class PluginForm : WeifenLuo.WinFormsUI.Docking.DockContent
        {
            public Guid cGuid;
            protected override string GetPersistString()
            {
                return cGuid.ToString();
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


        public void OcalcReload(PPLMain pPPLMain)
        {
            
            pPPLMain.ForceReloadDEP();
            pPPLMain.Rebuild3DDisplay();
            pPPLMain.UpdateCapacitySummary();
        }


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            progressBar1.Value = 50;
        }


    


        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
