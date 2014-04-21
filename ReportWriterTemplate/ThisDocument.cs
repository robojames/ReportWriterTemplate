using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace ReportWriterTemplate
{
    public partial class ThisDocument
    {

        #region Constant Declarations
        string ASTM_F1717 = "ASTM Standard F1717-13, \"Standard Test Methods for Spinal Implant Constructs in a Vertebrectomy Model.\"";
        string ASTM_F1714 = "ASTM F1714-96 Standard Guide, \"Gravimetric Wear Assessment of Prosthetic Hip-Designs in Simulator Devices, Annex A4.\"";
        string ASTM_F2077 = "ASTM Standard F2077-11, \"Test Methods for Intervertebral Body Fusion Devices.\"";
        string ASTM_F2267 = "ASTM Standard F2267-04, \"Standard Test Method for Measuring Load Induced Subsidence of Intervertebral Body Fusion Devices Under Static Axial Compression.\"";
        string ASTM_EXP = "ASTM Draft Standard F-04.25.02.02, \"Static Push-out Test Method for Intervertebral Body Fusion Devices,\" Draft #2 - August 29, 2000";

        string SystemCodeName = "NAMEOFSYSTEM";
        string JobCodeName = "123456";
        #endregion

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            
            MessageBox.Show("Welcome to James' Report Writer Template...\nUse with caution.  This shit is still in beta.");

        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.checkBox_TF.CheckedChanged += new System.EventHandler(this.checkBox_TF_CheckedChanged);
            this.checkBox_SF.CheckedChanged += new System.EventHandler(this.checkBox_SF_CheckedChanged);
            this.checkBox_F.CheckedChanged += new System.EventHandler(this.checkBox_F_CheckedChanged);
            this.checkBox_EX.CheckedChanged += new System.EventHandler(this.checkBox_EX_CheckedChanged);
            this.checkBox_TR.CheckedChanged += new System.EventHandler(this.checkBox_TR_CheckedChanged);
            this.checkBox_SC.CheckedChanged += new System.EventHandler(this.checkBox_SC_CheckedChanged);
            this.checkBox_AC.CheckedChanged += new System.EventHandler(this.checkBox_AC_CheckedChanged);
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);

        }

        #endregion

        /// <summary>
        /// Delegate which tells the program what to execute after the button is clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            Insert_System_Name();

            Insert_Job_Number();

            Generate_Objectives();

            Generate_Specimen_Description();

            Generate_Test_Specifications();



            MessageBox.Show("Completed Template Forming...");
        }


        public void Insert_System_Name()
        {
            var range = Application.ActiveDocument.Range();

            range.Find.Execute(FindText: SystemCodeName, Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: textBox_SystemName.Text, MatchCase: true);
            
        }

        public void Insert_Job_Number()
        {
            var range = Application.ActiveDocument.Range();

            range.Find.Execute(FindText: JobCodeName, Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: text_JobNumber.Text);
        }

        /// <summary>
        /// Generate the objectives of the test (mainly the bulleted list in Section 1.1)
        /// </summary>
        public void Generate_Objectives()
        {
            // Pulls range from bookmarked text
            Word.Range bullet_Tests = bookmark_BulletTests.Range;

            // Check which static tests are to be performed and insert the related bullet point
            if (checkBox_AC.Checked)
            {
                bullet_Tests.InsertAfter("Static testing in a load to failure mode in axial compression\n");
            }

            if (checkBox_SC.Checked)
            {
                bullet_Tests.InsertAfter("Static testing in a load to failure mode in compressive shear\n");
            }

            if (checkBox_TR.Checked)
            {
                bullet_Tests.InsertAfter("Static testing in a load to failure mode in torsion\n");
            }

            if (checkBox_EX.Checked)
            {
                bullet_Tests.InsertAfter("Static testing in a load to failure mode in expulsion\n");
            }


            // Check which dynamic tests are to be performed and insert the related bullet point
            if (checkBox_F.Checked)
            {
                bullet_Tests.InsertAfter("Cyclical axial compression testing to estimate the maximum run out load value at 5,000,000 cycles");
            }

            if (checkBox_SF.Checked)
            {
                bullet_Tests.InsertAfter("Cyclical compressive shear testing to estimate the maximum run out load value at 5,000,000 cycles");
            }

            if (checkBox_TF.Checked)
            {
                bullet_Tests.InsertAfter("Cyclical torsion testing to estimate the maximum run out load value at 5,000,000 cycles");
            }
                        
        }

        /// <summary>
        /// Generate the specimen description/orientation/build information
        /// </summary>
        public void Generate_Specimen_Description()
        {
            Word.Range Orientation_Information = bookmark_Orientation.Range;

            // If assembly was required, insert text based on whether we (ETC) or the customer did it.
            if (checkBox_CustAssembled.Checked)
            {
                // Insert customer assembled note, but first determine standard
                foreach (string item in List_TestedStandards.CheckedItems)
                {
                    // If testing standard is F1717, insert F1717 assembly information
                    if (item == "F1717")
                    {
                        Orientation_Information.InsertAfter("\n\nThe test specimens were received assembled by the customer with " + 
                        "UHMWPE OR polyacetal test blocks and with the bone OR pedicle set screws torqued to XXin-lbs.  The torque " +
                        "was verified to XXin-lbs on the bone OR pedicle set screws with customer provided instruments OR an ETC provided " +
                        "Proto 6106 Torque Screwdriver (Stanley Proto Industrial Tools, Covington, GA) OR an ETC provided Computorq3 " +
                        "Electronic Torque Wrench (CDI Torque Products, A Snap-on Company, City of Industry, CA).  A gap of XXmm between the " +
                        "screw housings and the test blocks was utilized to prevent impingement of pivoting features of the screw against the test block.\n");
                    }

                    // If testing standard is F2077, insert F2077 assembly information
                    if (item == "F2077")
                    {
                        Orientation_Information.InsertAfter("\n\nThe test specimens were received assembled by the customer with UHMWPE OR polyacetal OR " +
                            "stainless steel test blocks and with the bone screws torqued to XXin-lbs.  The torque was verified to XXin-lbs on the bone screws "+
                            "with the customer provided instruments OR an an ETC provided Cedar Digital Torque Screwdriver (Imada Inc., Northbrook, IL) OR an ETC " +
                            "provided Computorq3 Electronic Torque Wrench (CDI Torque Products, A Snap-on Company, City of Industry, CA).");
                    }
                }
            } // ETC Assembled Verbiage
            else if (checkBox_ETCAssembled.Checked)
            {
                // Insert ETC assembled verbiage
                foreach (string item in List_TestedStandards.CheckedItems)
                {
                    if (item == "F1717")
                    {
                        Orientation_Information.InsertAfter("\n\nThe test specimens were assembled in UHMWPE OR polyacetal test blocks " +
                            "per customer instructions.  The bone OR pedicle set screws were hand tightened OR tightened to a torque of " +
                            "XXin-lbs with the customer provided instruments OR an ETC provided Proto 6106 Torque Screwdriver (Stanley Proto " +
                            "Industrial Tools, Covington, GA) OR an ETC provided Computorq3 Electronic Torque Wrench (CDI Torque Products, A Snap-on " +
                            "Company, City of Industry, CA).  A gap of XXmm between the screw housings and the test blocks was utilized to prevent " +
                            "impingement of pivoting features of the screw against the test block.\n");
                    }

                    if (item == "F2077")
                    {
                        Orientation_Information.InsertAfter("\n\nThe test specimens were assembled in UHMWPE OR polyacetal OR stainless steel test blocks " +
                        "per customer instructions.  The bone screws were hand tightened OR tightened to a torque of XXin-lbs with the customer provided " +
                        "instruments OR an ETC provided Proto 6106 Torque Screwdriver (Stanley Proto Industrial Tools, Covington, GA) OR an ETC provided " +
                        "Computorq3 Electronic Torque Wrench (CDI Torque Products, A Snap-on Company, City of Industry, CA).");
                    }
                }
            }
            else
            {
                // Do nothing:
            }
                
            // Lastly, include the default orientation information for the given testing standard - This can probably be combined with the above section when refactored...
            foreach (string item in List_TestedStandards.CheckedItems)
            {
                if (item == "F2077")
                {
                    Orientation_Information.InsertAfter("\nEach specimen was oriented such that the laser etching was right side up.  An example of " +
                        "an untested specimen is shown in Figures XX - XX of Appendix A.  The test blocks are shown in Figures XX - XX of Appendix A " +
                        "and in Prints XX - XX of Appendix B.\n");
                }

                if (item == "F1717")
                {
                    Orientation_Information.InsertAfter("\nThe axial specimens were built with the transverse connector positioned at the vertical centerline " +
                        "of the rods, as shown in Figures XX - XX of Appendix A.  The torsion specimens were built without a transverse connector, as shown " +
                        "in Figures XX - XX of Appendix A.");

                    Orientation_Information.InsertAfter("\nOR\n An example of an untested specimen is shown in Figures XX - XX of Appendix A.\n\nThe test block print " +
                        "is shown in Print X of Appendix B.");
                }
            }
            
        }

        /// <summary>
        /// Fills out the testing specifications section, including if CDW was performed and which ASTM standards were followed
        /// </summary>
        public void Generate_Test_Specifications()
        {
            Word.Range Test_Specifications = bookmark_Standards_Protocols.Range;

            foreach (string item in List_TestedStandards.CheckedItems)
            {
                switch (item) {
                    case "Protocol":
                        Test_Specifications.InsertAfter("CUSTOMER PROTOCOL\n");
                        break;
                    case "F1717":
                        Test_Specifications.InsertAfter(ASTM_F1717);
                        break;
                    case "F2077":
                        Test_Specifications.InsertAfter(ASTM_F2077);
                        break;
                    
                }

            }
            
        }

        #region Check Box Delegates

        private void checkBox_AC_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_AC.Checked)
            {
                numericUpDown_AC.Enabled = true;
            }
            else
            {
                numericUpDown_AC.Enabled = false;
            }
        }

        private void checkBox_SC_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_SC.Checked)
            {
                numericUpDown_SC.Enabled = true;
            }
            else
            {
                numericUpDown_SC.Enabled = false;
            }
        }

        private void checkBox_TR_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_TR.Checked)
            {
                numericUpDown_TR.Enabled = true;
            }
            else
            {
                numericUpDown_TR.Enabled = false;
            }
        }

        private void checkBox_EX_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_EX.Checked)
            {
                numericUpDown_EX.Enabled = true;
            }
            else
            {
                numericUpDown_EX.Enabled = false;
            }
        }

        private void checkBox_F_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_F.Checked)
            {
                numericUpDown_F.Enabled = true;
            }
            else
            {
                numericUpDown_F.Enabled = false;
            }
        }

        private void checkBox_SF_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_SF.Checked)
            {
                numericUpDown_SF.Enabled = true;
            }
            else
            {
                numericUpDown_SF.Enabled = false;
            }
        }

        private void checkBox_TF_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_TF.Checked)
            {
                numericUpDown_TF.Enabled = true;
            }
            else
            {
                numericUpDown_TF.Enabled = false;
            }
        }

        #endregion


    }
}
