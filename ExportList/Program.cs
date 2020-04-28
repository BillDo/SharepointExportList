using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ExportList
{
    class Program

    {
        static void Main(string[] args)

        {
            using (SPSite site = new SPSite("http://sppets2/sites/eitq/"))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    foreach (SPList list in web.Lists)
                    {
                        DirectoryInfo dir = new DirectoryInfo(@"C:\List Data");
                        dir.Create();

                        FileInfo file = new FileInfo(@"C:\List Data\"+ list.Title + ".xls");
                        StreamWriter streamWriter = file.CreateText();

                        StringWriter stringWriter = new StringWriter();
                        HtmlTextWriter htmlTextWriter = new HtmlTextWriter(stringWriter);

                        Table tblData = GetListTableControl("http://sppets2/sites/eitq/Lists/" + list.Title, list.Title);
                        tblData.RenderControl(htmlTextWriter);
                        streamWriter.Write(stringWriter.ToString());


                        htmlTextWriter.Close();
                        streamWriter.Close();
                        stringWriter.Close();
                    }                    
                }
            }
        }

        public static Table GetListTableControl(string strListURL, string strListName)
        {
            Table tblListView = new Table();
            tblListView.ID = "_tblListView";
            tblListView.BorderStyle = BorderStyle.Solid;
            tblListView.BorderWidth = Unit.Pixel(1);

            using (SPSite site = new SPSite(strListURL.Trim()))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    SPList list = web.Lists[strListName.Trim()];

                    for (int i = 0; i < list.Fields.Count-1; i++)
                    {
                        tblListView.Rows.Add(new TableRow());

                        tblListView.Rows[i].Cells.Add(new TableCell());
                        tblListView.Rows[i].Cells[0].Text = list.Fields[i].Title;

                        tblListView.Rows[i].Cells.Add(new TableCell());
                        tblListView.Rows[i].Cells[1].Text = list.Fields[i].InternalName;

                        tblListView.Rows[i].Cells.Add(new TableCell());
                        tblListView.Rows[i].Cells[2].Text = list.Fields[i].TypeDisplayName;

                        tblListView.Rows[i].Cells.Add(new TableCell());
                        tblListView.Rows[i].Cells[3].Text = list.Fields[i].Required?"Yes":"No";

                        tblListView.Rows[i].Cells.Add(new TableCell());
                        tblListView.Rows[i].Cells[4].Text = list.Fields[i].Description;
                    }                    
                }
            }

            return tblListView;
        }
    }
}
