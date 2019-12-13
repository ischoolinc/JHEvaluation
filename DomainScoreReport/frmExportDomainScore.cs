using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using K12.Data;

namespace DomainScoreReport
{
    public partial class frmExportDomainScore : BaseForm
    {
        public frmExportDomainScore()
        {
            InitializeComponent();

            // Init SchoolYear Semester
            {
                int schoolYear = int.Parse(School.DefaultSchoolYear);
                cbxSchoolYear.Items.Add(schoolYear - 1);
                cbxSchoolYear.Items.Add(schoolYear);
                cbxSchoolYear.Items.Add(schoolYear + 1);
                cbxSchoolYear.SelectedIndex = 1;

                int semester = int.Parse(School.DefaultSemester);
                cbxSemester.Items.Add(1);
                cbxSemester.Items.Add(2);
                cbxSemester.SelectedIndex = semester - 1;
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {

        }

        private void btnLeave_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
