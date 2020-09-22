using System;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using System.Windows.Media.Imaging;

namespace Employee_System_Asst3
{
    public partial class Form1 : Form
    {
  
    OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Admin\Desktop\IT\IIBIT course\term2\C#\assignment\Employee_System-Asst3\Employee(Asst3).accdb");

        public Form1()
        {

            InitializeComponent();
            Page_Load();//function

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
        private void Page_Load()
        {
            //Retrieve data from access
            try
            {
                con.Open();
                String sql = "select * from employee";
                OleDbDataAdapter adapter = new OleDbDataAdapter(sql, con);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;


            }catch(Exception e)
            {
                MessageBox.Show(e + "");

            }
                        
            con.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //converting photo to binary data
             MemoryStream ms;
             byte[] photo_aray;
           
                //using MemoryStream:
                ms = new MemoryStream();
                pictureBox1.Image.Save(ms, ImageFormat.Jpeg);
           photo_aray = new byte[ms.Length];
                ms.Position = 0;
                ms.Read(photo_aray, 0, photo_aray.Length);
            try
            {             
           
            String fn = TextFN.Text;
            String ln = TextLN.Text;
            String age = TextAge.Text;
            String email = TextEmail.Text;
            String salary = TextSalary.Text;
            String note = richTextBox1.Text;
            String dpt = comboBox1.Text;
            String pos = TextPosition.Text;
            String DOB = dateTimePicker1.Text;
            String Hiredate = dateTimePicker2.Text;
            String img_location = textBox8.Text;
            //Set insert query
            string qry = "insert into employee([firstname],[lastname],[age],[email],[salary],[note],[department],[position]," +
                "[DOB],[HireDate],[img_location],[image]) values" +
                "('" + fn + "','" + ln + "','" + age + "','" + email + "','" + salary + "','" + note + "'," +
                "'" + dpt + "','" + pos + "','" + DOB + "','" + Hiredate + "','"+ img_location+"',@image)";



                //<------------------Store image  data------------------------>//
                con.Open();
                OleDbCommand connect = new OleDbCommand(qry, con);
                
                connect.Parameters.AddWithValue("@photo", photo_aray);

                connect.ExecuteNonQuery();
                MessageBox.Show("Data inserted successfully");


            }
            catch (Exception x)
            {
                MessageBox.Show(x+"");
            }
            finally
            {
                con.Close();
            }
            Page_Load();


        }





        private void Employee_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            String ImagePath;
            try
            {
                //use function OpenFileDialog to get file path
                OpenFileDialog dialog = new OpenFileDialog();

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    ImagePath = dialog.FileName; 
                    textBox8.Text = ImagePath;
                    pictureBox1.Image = Image.FromFile(ImagePath); //Display image on the box
                }

            }catch(Exception ex)
            {
                MessageBox.Show("An Error Occured","Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }



        

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string formatString = "dd-MM-yyyy";
           
            try
            {
                int Id = Convert.ToInt32(textBox1.Text.ToString()); //Selected ID
                String dateString = "dd-MM-yyyy"; //date format

                //SQL command
                string qry = "select * from [employee] where ID = @id";

                //<------------------Getting data------------------------>//
                con.Open();
                OleDbCommand connect = new OleDbCommand(qry,con);
                connect.Parameters.AddWithValue("@id", Id);
                OleDbDataReader reader = connect.ExecuteReader();

                while (reader.Read())
                {
                    TextID.Text = reader["ID"].ToString();
                    TextFN.Text = reader["firstname"].ToString();
                    TextLN.Text = reader["lastname"].ToString();
                    TextAge.Text = reader["age"].ToString();
                    TextEmail.Text = reader["email"].ToString();
                    TextSalary.Text = reader["salary"].ToString();
                    richTextBox1.Text = reader["note"].ToString();
                    comboBox1.Text = reader["department"].ToString();
                    TextPosition.Text = reader["position"].ToString();
                    textBox8.Text = reader["img_location"].ToString();
                    //Retrieve date as string from db and display on calendar format
                    string bd = reader["DOB"].ToString().Trim();
                    string hd = reader["HireDate"].ToString().Trim();
                    dateTimePicker1.Value = DateTime.ParseExact(bd, dateString, System.Globalization.CultureInfo.InvariantCulture);
                    dateTimePicker2.Value = DateTime.ParseExact(hd, dateString, System.Globalization.CultureInfo.InvariantCulture);

                    //Retrieve photo

                    byte[] image = (byte[])reader["image"];
                    if (image == null)
                    {
                        pictureBox1.Image = null;
                    }
                    else
                    {
                        MemoryStream mst = new MemoryStream(image);
                        pictureBox1.Image = System.Drawing.Image.FromStream(mst);
                    }
                }
            }
            catch (Exception x)
            {
                MessageBox.Show(x + "");
            }
            finally
            {
                con.Close();
            }
        }


  

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

       

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            TextID.Text = "";
            TextFN.Text = "";
            TextLN.Text = "";
            TextAge.Text = "";
            TextEmail.Text = "";
            TextSalary.Text = "";
            richTextBox1.Text = "";
            comboBox1.Text = "";
            TextPosition.Text = "";
            dateTimePicker1.Text = "";
            dateTimePicker2.Text = "";
            textBox8.Text = "";
            pictureBox1.Image = null;
        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            Page_Load();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            int Id = Convert.ToInt32(TextID.Text); //Selected ID
            //converting photo to binary data
             MemoryStream ms = new MemoryStream();
             byte[] photo_aray;

            //using MemoryStream:
             
             pictureBox1.Image.Save(ms,ImageFormat.Jpeg);
             photo_aray = new byte[ms.Length];
             ms.Position = 0;
             ms.Read(photo_aray, 0, photo_aray.Length);

                String fn = TextFN.Text.Trim();
                String ln = TextLN.Text.Trim();
                String age = TextAge.Text;
                String email = TextEmail.Text.Trim();
                String salary = TextSalary.Text.Trim();
                String note = richTextBox1.Text.Trim();
                String dpt = comboBox1.Text.Trim();
                String pos = TextPosition.Text.Trim();
                String DOB = dateTimePicker1.Text.Trim();
                String Hiredate = dateTimePicker2.Text.Trim();
                String img_location = textBox8.Text.Trim();

            //Set insert query
            
                       
             OleDbCommand cmd = new OleDbCommand();
             cmd.CommandType = CommandType.Text;
             cmd.CommandText = "UPDATE employee SET [firstname]=?,[lastname]=?,[age]=?,[email]=?,[salary]=?,[note]=?,[department]=?,[position]=?, [DOB]=?,[HireDate]=?,[img_location]=?,[Image]=? WHERE ID=?";


            try
            {

                if(pictureBox1.Image!=null){
              
              
                    cmd.Connection = con;
                    con.Open();
                    cmd.Parameters.AddWithValue("@firstname", fn);
                    cmd.Parameters.AddWithValue("@lastname", ln);
                    cmd.Parameters.AddWithValue("@age", age);
                    cmd.Parameters.AddWithValue("@email", email);
                    cmd.Parameters.AddWithValue("@salary", salary);
                    cmd.Parameters.AddWithValue("@note", note);
                    cmd.Parameters.AddWithValue("@dpt", dpt);
                    cmd.Parameters.AddWithValue("@pos", pos);
                    cmd.Parameters.AddWithValue("@dob", DOB);
                    cmd.Parameters.AddWithValue("@hire", Hiredate);
                    cmd.Parameters.AddWithValue("@path", img_location);
                    cmd.Parameters.AddWithValue("@photo", photo_aray);
                    cmd.Parameters.AddWithValue("@id", Id);
                    cmd.ExecuteNonQuery();
                    con.Close();
            } 
                
                Form1_Load(sender, e);
                MessageBox.Show("Data updated successfully");
                }

                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                        

            Page_Load();

                         
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int Id = Convert.ToInt32(TextID.Text); //Selected ID
            OleDbCommand cmd = new OleDbCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "DELETE FROM employee where ID=?";
            try
            {
                cmd.Connection = con;
                con.Open();
                cmd.Parameters.AddWithValue("@id", Id);
                cmd.ExecuteNonQuery();
                con.Close();

                Form1_Load(sender, e);
                MessageBox.Show("Data deleted");

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex+"");
            }

            Page_Load();

        }

       
        private void button7_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.ExitThread();
        }
    }
}

