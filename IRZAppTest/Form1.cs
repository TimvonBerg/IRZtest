using System;
using System.Windows.Forms;
using System.Net;
using System.Net.Http;
using Newtonsoft.Json;


namespace IRZAppTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (var client = new HttpClient(new HttpClientHandler { AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate })) //Gets or sets the type of decompression that is used
            {
                int pagesize = 30; //Number of responses to request

                string uri = "https://api.stackexchange.com/2.2/search?pagesize=" + pagesize + "&order=desc&sort=activity&intitle=beautiful&site=stackoverflow"; //Address of the request with variable pagesize

                HttpResponseMessage response = client.GetAsync(uri).Result; //Send a GET request to the specified Uri as an asynchronous operation

                response.EnsureSuccessStatusCode(); //Throws an exception if the IsSuccessStatusCode property for the HTTP response is false

                string result = response.Content.ReadAsStringAsync().Result; //Serialize the HTTP content to a string as an asynchronous operation

                Menu menu = JsonConvert.DeserializeObject<Menu>(result); //Calling the public class "Menu" to deserialize Json text

                object[,] DATAarray = new object[pagesize, 4]; //Set the array and the size of the array

                for (int a = 0; a < pagesize; a++) //Filling an array from Json
                {
                    DATAarray[a, 0] = menu.items[a].title;

                    DATAarray[a, 1] = menu.items[a].owner.display_name; //+ "\t" + menu.items[a].owner.link; //link to user accounts

                    if (menu.items[a].is_answered == true)
                    {
                        DATAarray[a, 2] = "Yes";
                    }
                    else
                    {
                        DATAarray[a, 2] = "No";
                    }

                    DATAarray[a, 3] = menu.items[a].link;
                }

                if (pagesize > 1) //Resizing dataGridView Table
                {
                    dataGridView1.Rows.Add(pagesize-1);
                }

                for (int i = 0; i < pagesize; i++)
                {
                    for (int j = 0; j < 4; j++)
                    {
                        dataGridView1.Rows[i].Cells[j].Value = DATAarray[i, j]; //Fill the dataGridView table from the array
                    }
                }  
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e) //Save the table and close the application
        {
            Save();

            Application.Exit();
        }

        private void Save() //Save function to excel
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);

            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            app.Visible = true; //if false, then excel will not be open

            worksheet = workbook.Sheets["Лист1"];

            worksheet = workbook.ActiveSheet;

            worksheet.Name = "api.stockchange.com";

            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
        }
    }
}
