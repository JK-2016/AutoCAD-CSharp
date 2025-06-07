using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Civil.Software
{
    public partial class CC_Wall_Input: Form
    {
        private static double last_St = 0;
        private static double last_H = 0;
        private static double last_Ft = 0;
        private static double last_FB = 0;
        private static double last_RB = 0;
        private static double last_FOF = 0;
        private static double last_ROF = 0;
        private static double last_FBH = 0;
        private static double last_RBH = 0;
        private static int last_rebar_m = 0;
        private static int last_spacing_m = 0;
        private static int last_rebar_d = 0;
        private static int last_spacing_d = 0;

        public double St_o { get; private set; }
        public double H_o { get; private set; }
        public double Ft_o { get; private set; }
        public double FB_o { get; private set; }
        public double RB_o { get; private set; }
        public double FOF_o { get; private set; }
        public double ROF_o { get; private set; }

        public double FBH_o { get; private set; }
        public double RBH_o { get; private set; }

        public int rebar_m_o { get; private set; }
        public int spacing_m_o { get; private set; }

        public int rebar_d_o { get; private set; }
        public int spacing_d_o { get; private set; }


        public CC_Wall_Input()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
            St.Text = last_St.ToString();
            H.Text = last_H.ToString();
            Ft.Text = last_Ft.ToString();
            FB.Text = last_FB.ToString();
            RB.Text = last_RB.ToString();
            FOF.Text = last_FOF.ToString();
            ROF.Text = last_ROF.ToString();
            FBH.Text = last_FBH.ToString();
            RBH.Text = last_RBH.ToString();
            bardia_m.Text = last_rebar_m.ToString();
            spacing_m.Text = last_spacing_m.ToString();
            bardia_d.Text = last_rebar_d.ToString();
            spacing_d.Text = last_spacing_d.ToString();
        }

        private void Draw_Click(object sender, EventArgs e)
        {
            bool valid = true;
            double st = 0, h = 0, ft = 0, fb = 0, rb = 0, fof = 0, rof = 0, fbh=0,rbh =0;
            int rebar_m=10, spac_m = 0, rebar_d = 10, spac_d = 0;
            valid = valid && double.TryParse(St.Text, out  st);
            valid = valid && double.TryParse(H.Text, out  h);
            valid = valid && double.TryParse(Ft.Text, out  ft);
            valid = valid && double.TryParse(FB.Text, out  fb);
            valid = valid && double.TryParse(RB.Text, out  rb);
            valid = valid && double.TryParse(FOF.Text, out  fof);
            valid = valid && double.TryParse(ROF.Text, out  rof);
            valid = valid && double.TryParse(FBH.Text, out fbh);
            valid = valid && double.TryParse(RBH.Text, out rbh);
            if(Yes.Checked)
            {
                valid = valid && int.TryParse(bardia_m.Text, out rebar_m);
                valid = valid && int.TryParse(spacing_m.Text, out spac_m);
                valid = valid && int.TryParse(bardia_d.Text, out rebar_d);
                valid = valid && int.TryParse(spacing_d.Text, out spac_d);

            }

            if (valid)
            {
                St_o = st;
                H_o = h;
                Ft_o = ft;
                FB_o = fb;
                RB_o = rb;
                FOF_o = fof;
                ROF_o = rof;
                FBH_o = fbh;
                RBH_o = rbh;
                rebar_m_o = rebar_m;
                spacing_m_o = spac_m;
                rebar_d_o = rebar_d;
                spacing_d_o = spac_d;
                //Save the values for next time
                last_St = St_o;
                last_H = H_o;
                last_Ft = Ft_o;
                last_FB = FB_o;
                last_RB = RB_o;
                last_FOF = FOF_o;
                last_ROF = ROF_o;
                last_FBH = FBH_o;
                last_RBH = RBH_o;
                last_rebar_d = rebar_d;
                last_spacing_d = spac_d;
                last_rebar_m = rebar_m;
                last_spacing_m = spac_m;

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("Please enter valid numeric values.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

     

        
    }
}
