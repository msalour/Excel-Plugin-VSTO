using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;
using static Excel_Plugin.RibbonCustom;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System.Linq;

namespace Excel_Plugin
{
    public partial class RibbonCustom
    {
        UserControl1 ctrl;
        Form1 form;
        private Microsoft.Office.Tools.CustomTaskPane taskPane;

        private async void RibbonCustom_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                // Set the welcome label
                Lbl_welcome.Label = "Hello " + Globals.ThisAddIn.Application.UserName + " (" + Environment.UserName + ")";
                checkBox1.Checked = true;
                checkBox2.Checked = false;
                toggleButton1.Checked = true;

                // Call the functions to populate the dropdown and combo box with API values
                await PopulateDropDownWithAPIValuesAsync();
                await PopulateComboBoxWithAPIValuesAsync();
                // Add this line to fill the spreadsheet with API data
                await FillSpreadsheetWithAPIData();

            }
            catch (Exception ex)
            {

            }
        }


        private async Task FillSpreadsheetWithAPIData()
        {
            try
            {
                string apiUrl = "https://jsonplaceholder.typicode.com/todos?_limit=5";
                using (var httpClient = new HttpClient())
                {
                    HttpResponseMessage response = await httpClient.GetAsync(apiUrl);

                    if (response.IsSuccessStatusCode)
                    {
                        string json = await response.Content.ReadAsStringAsync();
                        JArray data = JArray.Parse(json);

                        // Check if there is an active worksheet, otherwise create a new workbook and worksheet
                        Worksheet worksheet = GetActiveWorksheetOrCreateNew();

                        // Populate data
                        PopulateWorksheetWithData(worksheet, data);
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions
            }
        }

        private Worksheet GetActiveWorksheetOrCreateNew()
        {
            if (Globals.ThisAddIn.Application.ActiveSheet is Worksheet activeWorksheet)
            {
                return activeWorksheet;
            }
            else
            {
                Workbook newWorkbook = Globals.ThisAddIn.Application.Workbooks.Add();
                return (Worksheet)newWorkbook.Worksheets[1];
            }
        }

        private void PopulateWorksheetWithData(Worksheet worksheet, JArray data)
        {
            var startCell = worksheet.Cells[1, 1];
            JToken firstObject = data.Children().First();

            if (firstObject != null)
            {
                int col = 1;
                foreach (var property in firstObject.Children<JProperty>())
                {
                    worksheet.Cells[1, col] = property.Name;
                    col++;
                }
            }

            int row = 2;
            foreach (var item in data.Children())
            {
                int col = 1;
                foreach (var property in item.Children<JProperty>())
                {
                    worksheet.Cells[row, col] = property.Value.ToString();
                    col++;
                }
                row++;
            }
        }
        private async Task PopulateDropDownWithAPIValuesAsync()
        {
            using (var httpClient = new HttpClient())
            {
                try
                {
                    string apiUrl = "https://jsonplaceholder.typicode.com/posts?_limit=5";
                    HttpResponseMessage response = await httpClient.GetAsync(apiUrl);

                    if (response.IsSuccessStatusCode)
                    {
                        string json = await response.Content.ReadAsStringAsync();
                        List<Post> posts = JsonConvert.DeserializeObject<List<Post>>(json);

                        // Clear existing items in the dropdown
                        dropDown1.Items.Clear();

                        foreach (var post in posts)
                        {
                            var item = Factory.CreateRibbonDropDownItem();
                            item.Label = post.Title;
                            item.Tag = post.Id;
                            dropDown1.Items.Add(item);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Handle exceptions
                }
            }
        }

        public class Post
        {
            public int Id { get; set; }
            public string Title { get; set; }
        }

        private async Task PopulateComboBoxWithAPIValuesAsync()
        {
            using (var httpClient = new HttpClient())
            {
                try
                {
                    string apiUrl = "https://jsonplaceholder.typicode.com/todos?_limit=5";
                    HttpResponseMessage response = await httpClient.GetAsync(apiUrl);

                    if (response.IsSuccessStatusCode)
                    {
                        string json = await response.Content.ReadAsStringAsync();
                        List<Todo> todos = JsonConvert.DeserializeObject<List<Todo>>(json);

                        // Clear existing items in the ComboBox
                        comboBox1.Items.Clear();

                        foreach (var todo in todos)
                        {
                            var item = Factory.CreateRibbonDropDownItem();
                            item.Label = todo.Title;
                            item.Tag = todo.Id;
                            comboBox1.Items.Add(item);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Handle exceptions
                }
            }
        }

        public class Todo
        {
            public int Id { get; set; }
            public string Title { get; set; }
        }

        private void Btn_TaskPane_Click(object sender, RibbonControlEventArgs e)
        {
            ctrl = new UserControl1();
            taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(ctrl, "VSTO TASK PANE");
            taskPane.Width = 300;
            taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
            taskPane.Visible = true;
        }

        private void Btn_WinForm_Click(object sender, RibbonControlEventArgs e)
        {
            form = new Form1();
            form.Show();
        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            UpdateEditBox3Value();
        }

        private void editBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {
            UpdateEditBox3Value();
        }

        private void UpdateEditBox3Value()
        {
            try
            {
                int value1 = int.Parse(editBox1.Text);
                int value2 = int.Parse(editBox2.Text);
                editBox3.Text = (value1 + value2).ToString();
            }
            catch (FormatException)
            {
                // Handle format exception
            }
        }
    }
}