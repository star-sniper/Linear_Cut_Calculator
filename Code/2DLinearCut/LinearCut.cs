using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Windows.Forms;

namespace _2DLinearCut
{
    public partial class LinearCut : Form
    {
        double[] drop_length;
        int[] drop_qty;
        int drop_ctr = 0;
        int No_Of_Planks = 0;
        double saw_width = 0;
        double plankLength = 0;


        public LinearCut()
        {
            InitializeComponent();
        }

        //SHOW RESULT BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(textBox1.Text))
            {
                bool val = Double.TryParse(textBox1.Text, out plankLength);
                if (val != true || plankLength == 0)
                {
                    MessageBox.Show("Lenght should be greater than 0!");
                }
                else if (val == true && plankLength > 0)
                {
                    getGridValues();
                    fix_arrays();
                    calculate_planks();
                    string desctination = "C:\\Users\\" + Environment.UserName + "\\Desktop\\INS-DUMP\\Cut Ins.pdf";
                    System.Diagnostics.Process.Start(desctination);
                }
            }
            else
            {
                MessageBox.Show("Length can't be empty!");
            }
        }

        //PLANK LENTH LEAVE
        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(textBox1.Text))
            {
                bool val = Double.TryParse(textBox1.Text, out plankLength);
                if (val != true || plankLength == 0)
                {
                    MessageBox.Show("Lenght should be greater than 0!");
                }
            }
            else
            {
                MessageBox.Show("Length can't be empty!");
            }
        }

        //SAW WIDTH LEAVE
        private void textBox2_Leave(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(textBox2.Text))
            {
                bool val = Double.TryParse(textBox2.Text, out saw_width);
                if (val != true || saw_width == 0)
                {
                    MessageBox.Show("Saw Width should be greater than 0!");
                }
            }
            else
            {
                MessageBox.Show("Saw width can't be empty!");
            }
        }

        //PLANK CLASS FUNCTION
        List<Plank> CalculateCuts(List<double> desired, double plank_len)
        {
            List<double> PossibleLengths = new List<double> { };

            double Olength = plank_len;
            PossibleLengths.Add(Olength);
            var planks = new List<Plank>(); //Buffer list
                                            //go through cuts
            foreach (var i in desired)
            {
                //if no eligible planks can be found
                if (!planks.Any(plank => plank.lengthRem() >= i))
                {
                    //make a plank
                    planks.Add(new Plank(PossibleLengths.Max()));
                    No_Of_Planks++;
                }

                //cut where possible
                foreach (var plank in planks.Where(plank => plank.lengthRem() >= i))
                {
                    plank.Cut(i);
                    break;
                }

            }

            //cut down on waste by minimising length of plank
            foreach (var plank in planks)
            {
                double newLength = plank.plankLength;
                foreach (double possibleLength in PossibleLengths)
                {
                    if (possibleLength != plank.plankLength && plank.plankLength - plank.lengthRem() <= possibleLength)
                    {
                        newLength = possibleLength;
                        break;
                    }
                }
                plank.plankLength = newLength;
            }
            return planks;
        }

        //PLANK CLASS
        private void calculate_planks()
        {
            DataTable dt = new DataTable();
            
            int dt_rows = 0;
            dt.Columns.Add();
            dt.Columns.Add();

            for (int i = 0; i < drop_ctr; i++)
            {
                if (drop_length[i] > 0)
                {
                    DataRow dr = dt.NewRow();
                    dt.Rows.Add(dr);
                    dt.Rows[dt_rows][0] = drop_length[i];
                    dt.Rows[dt_rows][1] = drop_qty[i];
                    dt_rows++;
                }
            }

            No_Of_Planks = 0;

            int d = dt.Rows.Count;
            List<double> DesiredLengths = new List<double> { };

            double[] size_req = new double[d];
            double[] rods_req = new double[d];

            for (int f = 0; f < d; f++)
            {
                size_req[f] = Convert.ToDouble(dt.Rows[f][1]);
                rods_req[f] = Convert.ToDouble(dt.Rows[f][0]);

                //The cuts to be made
                for (int j = 0; j < size_req[f]; j++)
                {
                    DesiredLengths.Add(rods_req[f]);
                }
            }

            //Perform some basic optimisations
            DesiredLengths.Sort();
            DesiredLengths.Reverse();

            //Cut!
            var planks = CalculateCuts(DesiredLengths, plankLength);
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                dt.Rows.RemoveAt(i);
            }
            for (int i = dt.Columns.Count - 1; i >= 0; i--)
            {
                dt.Columns.RemoveAt(i);
            }
            List<int> Cuts = new List<int> { };
            string header = "", instructions = "";
            header = "CUTTING INSTRUCTION";
            header += "\nWILL NEED (" + No_Of_Planks + ") LENGTHS OF for cutting these items:";
            try
            {
                int pattern = 1;
                int count = 0;
                List<double> a = new List<double>();
                double b = 0f;
                foreach (var i in planks)
                {
                    if (!(a == i.Cuts || b == i.lengthRem()))
                    {
                        if (a.Any())
                        {
                            instructions += "\nPattern (" + pattern + ") from " + i.plankLength / 12 + "'-0\" lengths." + " Needed: " + (count == 1 ? count + " length" : count + " lengths") + ". Unused " + b + "\" each length\n";
                            instructions += count_ins(a, saw_width);
                            pattern++;
                        }
                        count = 0;
                    }
                    a = i.Cuts;
                    b = i.lengthRem();
                    count++;
                }
                instructions += "\nPattern (" + pattern + ") from " + planks[0].plankLength / 12 + "'-0\" lengths." + " Needed: " + (count == 1 ? count + " length" : count + " lengths") + ". Unused " + b + "\" each length\n";
                instructions += count_ins(a, saw_width);
            }
            catch (Exception exe)
            {
                MessageBox.Show("Exe " + exe);
            }
            create_instructions(ref header, ref instructions);
        }

        //GETTING VALUES FROM GRID TO TABLE
        private void getGridValues()
        {
            drop_ctr = 0;
            drop_length = new double[dataGridView1.Rows.Count - 1];
            drop_qty = new int[dataGridView1.Rows.Count - 1];

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value != null && dataGridView1.Rows[i].Cells[1].Value != null)
                {
                    string length = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    string qty = dataGridView1.Rows[i].Cells[1].Value.ToString();

                    bool val = Double.TryParse(length, out double len_each);
                    bool val2 = Int32.TryParse(qty, out int qty_each);

                    if (len_each != 0 && qty_each != 0)
                    {
                        drop_length[drop_ctr] = len_each;
                        drop_qty[drop_ctr] = qty_each;
                        drop_ctr++;
                    }
                }
            }
        }

        //FIXES THE DROP LENGTH AND DROP QTY ARRAY, TOTALS THE QUANTITES OF SAME ELEMENT OCCURING MORE THAN ONCE THEREBY MAKING IT A SINGLE ELEMENT
        private void fix_arrays()
        {
            int arraySize = dataGridView1.Rows.Count - 1;
            double[] temp_drop_length = new double[arraySize];
            int[] temp_drop_qty = new int[arraySize];

            int temp_ctr = 0;
            int count = 0;
            double get_first = 0;
            int qty = 0;

            for (int i = 0; i < drop_ctr; i++)
            {
                get_first = drop_length[i];
                for (int xy = i; xy < drop_ctr; xy++)
                {
                    if (get_first > 0 && get_first == drop_length[xy])
                    {
                        count++;
                        qty += drop_qty[xy];
                        drop_length[xy] = 0;
                        drop_qty[xy] = 0;
                    }
                }
                if (count >= 1)
                {
                    temp_drop_length[temp_ctr] = get_first;
                    temp_drop_qty[temp_ctr] = qty;
                    temp_ctr++;
                    count = 0;
                    qty = 0;
                }
            }

            drop_ctr = temp_ctr;
            drop_length = new Double[drop_ctr];
            drop_qty = new int[drop_ctr];
            for (int i = 0; i < temp_ctr; i++)
            {
                drop_length[i] = temp_drop_length[i];
                drop_qty[i] = temp_drop_qty[i];
            }
        }

        //COUNTS OCCURENCES OF AN ITEM IN THE LIST, USED WHILE GENRATING THE DOC FILE
        private string count_ins(List<double> a, double sawwidth)
        {
            string final = "";
            int count = 0;
            double[] arr = a.ToArray();

            for (int i = 0; i < arr.Length; i++)
            {
                arr[i] = arr[i] - sawwidth;
            }

            for (int i = 0; i < arr.Length; i++)
            {
                double initial = arr[i];
                if (initial > 0)
                {
                    for (int j = i; j < arr.Length; j++)
                    {
                        if (initial == arr[j])
                        {
                            arr[j] = 0;
                            count++;
                        }
                    }
                    final += "\nCut Length: " + initial + "\" Number cut from each stock piece: " + (count) + "\n";
                    count = 0;
                }
            }

            return final;
        }

        //CREATES A PDF WITH CUTTING INSTRUCTIONS
        private void create_instructions(ref string header, ref string instruc)
        {
            DataSet ds = new DataSet();
            DataTable dx = new DataTable();
            dx.Columns.Add("BOM Items");
            dx.Columns.Add("Quantity");

            int dt_counter = 0;

            //       connection.Open();
            for (int i = 0; i < drop_ctr; i++)
            {
                DataRow dr = dx.NewRow();
                dx.Rows.Add(dr);
                dx.Rows[dt_counter][0] = drop_length[i];
                dx.Rows[dt_counter][1] = drop_qty[i];
                dt_counter++;
            }
            //       connection.Close();

            ds.Tables.Add(dx);
            ds.WriteXmlSchema("po_instruction.xml");
            PO_instructions cs = new PO_instructions();
            cs.SetDataSource(ds);
            TextObject t = (TextObject)cs.ReportDefinition.Sections[0].ReportObjects["header"];
            t.Text += header;
            TextObject ins = (TextObject)cs.ReportDefinition.Sections[2].ReportObjects["patterns"];
            ins.Text = instruc;
            crystalReportViewer1.ReportSource = cs;
            if (!System.IO.Directory.Exists("C:\\Users\\" + Environment.UserName + "\\Desktop\\INS-DUMP"))
            {
                System.IO.Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Desktop\\INS-DUMP");
            }
           // cs.ExportToDisk(ExportFormatType.PortableDocFormat, "C:\\Users\\" + Environment.UserName + "\\Desktop\\INS-DUMP\\Cut Ins(" + pageno + ").pdf");
            cs.ExportToDisk(ExportFormatType.PortableDocFormat, "C:\\Users\\" + Environment.UserName + "\\Desktop\\INS-DUMP\\Cut Ins.pdf");
        }

        //RESETTING THE FORM
        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            textBox1.Text = "";
            textBox2.Text = "";
        }

        //choose excel file
        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string cmd_cnt = @"Provider = Microsoft.ACE.OleDb.12.0; Data Source =" + openFileDialog1.FileName + ";Extended Properties=\"Excel 12.0;HDR=Yes;\";";
                DataTable dt = new DataTable();
                string cmd = "SELECT * FROM [Sheet1$]";
                OleDbConnection ole = new OleDbConnection(cmd_cnt);

                OleDbDataAdapter mydata = new OleDbDataAdapter(cmd, ole);
                try
                {
                    mydata.Fill(dt);
                }
                catch (Exception err)
                {
                    MessageBox.Show("Error" + err);
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i][0] != null && dt.Rows[i][0].ToString() != ""
                        && dt.Rows[i][1] != null && dt.Rows[i][1].ToString() != "")
                    {
                        string length = dt.Rows[i][0].ToString();
                        string qty = dt.Rows[i][1].ToString();

                        bool val = Double.TryParse(length, out double len_each);
                        bool val2 = Int32.TryParse(qty, out int qty_each);

                        if (val && val2)
                        {
                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[i].Cells[0].Value = dt.Rows[i][0].ToString();
                            dataGridView1.Rows[i].Cells[1].Value = dt.Rows[i][1].ToString();
                        }
                    }
                }
            }
        }
    }
}
