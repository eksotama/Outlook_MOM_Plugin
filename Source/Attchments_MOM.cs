using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CalScanner
{
    public partial class Attchments_MOM_Form : Form
    {
        public Attchments_MOM_Form()
        {
            InitializeComponent();
        }
        public void loadData(String fullName)
        {            
           dataGridView1.Rows.Add(fullName.Substring(fullName.LastIndexOf("\\")+1).Trim(),fullName.Trim());
        }
    }
}
