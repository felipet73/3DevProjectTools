﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Json;
using Newtonsoft.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using EO.WebBrowser;
using Biblioteca;
using RestSharp;
//using Microsoft.Office.Interop.MSProject;


using MSProject = Microsoft.Office.Interop.MSProject;
using Office = Microsoft.Office.Core;


//using Project = Microsoft.Office.Interop.MSProject;



namespace ProjectAddIn1
{
    public partial class UserControl1 : UserControl
    {
        
        public UserControl1()
        {
            EO.WebBrowser.Runtime.AddLicense(
           "3a5rp7PD27FrmaQHEPGs4PP/6KFrqKax2r1GgaSxy5916u34GeCt7Pb26bSG" +
           "prT6AO5p3Nfh5LRw4rrqHut659XO6Lto6u34GeCt7Pb26YxDs7P9FOKe5ff2" +
           "6YxDdePt9BDtrNzCnrWfWZekzRfonNzyBBDInbW1yQKzbam2xvGvcau0weKv" +
           "fLOz/RTinuX39vTjd4SOscufWbPw+g7kp+rp9unMmuX59hefi9jx+h3ks7Oz" +
           "/RTinuX39hC9RoGkscufddjw/Rr2d4SOscufWZekscu7mtvosR/4qdzBs/DO" +
           "Z7rsAxrsnpmkBxDxrODz/+iha6iywc2faLWRm8ufWZfAwAzrpeb7z7iJWZek" +
           "sefuq9vpA/Ttn+ak9QzznrSmyNqxaaa2wd2wW5f3Bg3EseftAxDyeuvBs+I=");
            InitializeComponent();
            webView1.JSExtInvoke += new JSExtInvokeHandler(WebView_JSExtInvoke);
        }

        void WebView_JSExtInvoke(object sender, JSExtInvokeArgs e)
        {
            switch (e.FunctionName)
            {
                case "demoAbout":
                    /*string browserEngine = e.Arguments[0] as string;
                    string url = e.Arguments[1] as string;
                    MessageBox.Show("Browser Engine: " + browserEngine + ", Url:" + url);*/

                    /*Excel.Workbook libro = Globals.ThisAddIn.Application.ActiveWorkbook;
                    Excel.Worksheet hoja = libro.Worksheets[1];*/
                    //hoja.Cells[2, 1] = e.Arguments[2];
                    string ID = e.Arguments[0] as string;
                    JsonValue data = System.Json.JsonObject.Parse(e.Arguments[2] as string);

                    string Nombre = e.Arguments[1] as string;


                    //hoja.Cells[2, 1] = (string)ID;

                    for (int X = 0; X < data.Count; X++)
                    {
                        if ((string)data[X]["attributeName"] == "Category") {
                            Nombre = Nombre + " " + data[X]["displayValue"];
                        }
                        if ((string)data[X]["attributeName"] == "Type Name")
                        {
                            Nombre = Nombre + " Tipo: (" + data[X]["displayValue"] + ")" ; 
                        }

                        //hoja.Cells[3 + X, 2] = (string)data[X]["displayName"];
                        //hoja.Cells[3 + X, 3] = data[X]["displayValue"];
                    }
                    //string name = (string)data["name"];
                    //string version = (string)data["version"];

                    MSProject.Project prj;
                    prj = Globals.ThisAddIn.Application.ActiveProject;
                    //MSProject.Task tarea;
                    //tarea.Name = Nombre;
                    //tarea.Text1 = e.Arguments[2] as string;
                    //prj.Tasks.Add(Nombre);
                    MSProject.Task tarea  = prj.Tasks.Add( Nombre );
                    tarea.Text1 = ID;
                    break;
            }
        }

        private void groupPanel1_Click(object sender, EventArgs e)
        {

        }
        public void ejecutar(string ID) {
            webView1.EvalScript("highlightRevit('" + ID + "');");
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {
            //webControl1.WebView.LoadUrl("https://www.facebook.com/");        
            //webControl1.WebView.LoadUrl("http://localhost:3001/");
            //"file:///HTML/Viewer.html?URN=dXJuOmFkc2sud2lwcHJvZDpmcy5maWxlOnZmLkpNR2ZnSFh2VFVHNmFRNkNBWTVOcWc/dmVyc2lvbj0x&Token=eyJhbGciOiJSUzI1NiIsImtpZCI6IlU3c0dGRldUTzlBekNhSzBqZURRM2dQZXBURVdWN2VhIn0.eyJzY29wZSI6WyJkYXRhOnJlYWQiLCJkYXRhOndyaXRlIiwiZGF0YTpjcmVhdGUiLCJkYXRhOnNlYXJjaCIsImJ1Y2tldDpjcmVhdGUiLCJidWNrZXQ6cmVhZCIsImJ1Y2tldDp1cGRhdGUiLCJidWNrZXQ6ZGVsZXRlIl0sImNsaWVudF9pZCI6IkxybjZvcUxud3BDQmQ4R1MwTHVpbUd4NVNIT05ZdzRiIiwiYXVkIjoiaHR0cHM6Ly9hdXRvZGVzay5jb20vYXVkL2Fqd3RleHA2MCIsImp0aSI6IkdmRzNaVmxtWDlzWkM4a0lrV2F5UjFrQXJrekdXVHQ3OXdaczlzckoycTZBZ0h1d1hBU3YzanRYcDlFSlJUZ04iLCJleHAiOjE2NDEwODQzNzV9.Oq7XE3bDoG1m1T9ZBUcRSCYZy_FJPv7jwho9uypdIv3aHgM7DVONQLIQ5a3PdZkdFzxImO5C4x_RRJyFCaJwXDsvcgW8jR_mTzNqbNQU1kfIeDrMMT5-bc0Tgt1galK9NMBXrM9eeUsXAcKTF_UY9blnnLY8ef9_8NsRRuVsjSdaZ7UK3H6qx1lWIChCuGqgi3ODwgJuv9fyv3GO0a1ULIDVLTuOwp5NowqsAWXtf7BpB0veEpnaTwPsvLU72hxJV3ar-9QJ8gJI8aZ7F-055asuPVvbTKRtWQIQGwGSOUHwePX16R-elEUP3QvjpwG8P9FUMjPZQryuCP3PmPw6ng"

            //webControl1.WebView.LoadUrl(ViewerURN("urn:adsk.wipprod:fs.file:vf.JMGfgHXvTUG6aQ6CAY5Nqg?version=1"),"");
            //webControl1.WebView.LoadUrl(@"file:///C:/HTML/Viewer.html?URN=dXJuOmFkc2sud2lwcHJvZDpmcy5maWxlOnZmLkpNR2ZnSFh2VFVHNmFRNkNBWTVOcWc/dmVyc2lvbj0x&Token=eyJhbGciOiJSUzI1NiIsImtpZCI6IlU3c0dGRldUTzlBekNhSzBqZURRM2dQZXBURVdWN2VhIn0.eyJzY29wZSI6WyJkYXRhOnJlYWQiLCJkYXRhOndyaXRlIiwiZGF0YTpjcmVhdGUiLCJkYXRhOnNlYXJjaCIsImJ1Y2tldDpjcmVhdGUiLCJidWNrZXQ6cmVhZCIsImJ1Y2tldDp1cGRhdGUiLCJidWNrZXQ6ZGVsZXRlIl0sImNsaWVudF9pZCI6IkxybjZvcUxud3BDQmQ4R1MwTHVpbUd4NVNIT05ZdzRiIiwiYXVkIjoiaHR0cHM6Ly9hdXRvZGVzay5jb20vYXVkL2Fqd3RleHA2MCIsImp0aSI6IkdmRzNaVmxtWDlzWkM4a0lrV2F5UjFrQXJrekdXVHQ3OXdaczlzckoycTZBZ0h1d1hBU3YzanRYcDlFSlJUZ04iLCJleHAiOjE2NDEwODQzNzV9.Oq7XE3bDoG1m1T9ZBUcRSCYZy_FJPv7jwho9uypdIv3aHgM7DVONQLIQ5a3PdZkdFzxImO5C4x_RRJyFCaJwXDsvcgW8jR_mTzNqbNQU1kfIeDrMMT5-bc0Tgt1galK9NMBXrM9eeUsXAcKTF_UY9blnnLY8ef9_8NsRRuVsjSdaZ7UK3H6qx1lWIChCuGqgi3ODwgJuv9fyv3GO0a1ULIDVLTuOwp5NowqsAWXtf7BpB0veEpnaTwPsvLU72hxJV3ar-9QJ8gJI8aZ7F-055asuPVvbTKRtWQIQGwGSOUHwePX16R-elEUP3QvjpwG8P9FUMjPZQryuCP3PmPw6ng");
            webControl1.WebView.LoadUrl(@"file:///c:/HTML3/Viewer.html?URN=dXJuOmFkc2sud2lwcHJvZDpmcy5maWxlOnZmLmJUQlZFNDl1VFlDcm8tT3gxRVA1bkE/dmVyc2lvbj0x&Token=eyJhbGciOiJSUzI1NiIsImtpZCI6IlU3c0dGRldUTzlBekNhSzBqZURRM2dQZXBURVdWN2VhIn0.eyJzY29wZSI6WyJkYXRhOnJlYWQiLCJkYXRhOndyaXRlIiwiZGF0YTpjcmVhdGUiLCJkYXRhOnNlYXJjaCIsImJ1Y2tldDpjcmVhdGUiLCJidWNrZXQ6cmVhZCIsImJ1Y2tldDp1cGRhdGUiLCJidWNrZXQ6ZGVsZXRlIl0sImNsaWVudF9pZCI6IkxybjZvcUxud3BDQmQ4R1MwTHVpbUd4NVNIT05ZdzRiIiwiYXVkIjoiaHR0cHM6Ly9hdXRvZGVzay5jb20vYXVkL2Fqd3RleHA2MCIsImp0aSI6InVkT1dSRFV3TGZGSXl5VUFHbWhRUkpaZHlxWmZvSXJnUlFFcmJFOHRqZ1A3N1h6UmRSZFU0ZjVramw2NkpnWlIiLCJleHAiOjE2NDMyOTc2Nzh9.Uj1PnIL-LounXAnHM_LnGJns9EqBQP6nWC7LJ0O5CEt7Xal_p6eG79CIPdBQvgroESby27OV10ZUHP2PNw5NyQscc4HhN8Edc040LOMKcacyqP9HX1HQ2_k66JvD4FMNRw4QOWoQtH_ToytbHwBy8awJ7SaUbvOkOtnje3qRzjgNB-B2FCrKLX9J_AFaEHui6yPzXpz9ouWo8bM-oFQfVZq-cFQMD48_DhAhXEWV30af6sg3KLc-4QKHrAta2_GMtlWfASjps57N-3Fn0FL5hXpHraHbKv5lvZuyAI1QGLM-rcCsrd1Fs6jiSNBk8kT-rQrGW0FWWPVrowzcoE-iCA");
            EO.WebBrowser.Runtime.AddLicense(
            "3a5rp7PD27FrmaQHEPGs4PP/6KFrqKax2r1GgaSxy5916u34GeCt7Pb26bSG" +
            "prT6AO5p3Nfh5LRw4rrqHut659XO6Lto6u34GeCt7Pb26YxDs7P9FOKe5ff2" +
            "6YxDdePt9BDtrNzCnrWfWZekzRfonNzyBBDInbW1yQKzbam2xvGvcau0weKv" +
            "fLOz/RTinuX39vTjd4SOscufWbPw+g7kp+rp9unMmuX59hefi9jx+h3ks7Oz" +
            "/RTinuX39hC9RoGkscufddjw/Rr2d4SOscufWZekscu7mtvosR/4qdzBs/DO" +
            "Z7rsAxrsnpmkBxDxrODz/+iha6iywc2faLWRm8ufWZfAwAzrpeb7z7iJWZek" +
            "sefuq9vpA/Ttn+ak9QzznrSmyNqxaaa2wd2wW5f3Bg3EseftAxDyeuvBs+I=");
        }


        string Base64Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            var ddd = Convert.ToBase64String(plainTextBytes);
            return Convert.ToBase64String(plainTextBytes);
        }

        string ViewerURN(string urn, string viewableId)
        {
            string respuesta = string.Empty;
            var curiosidad = Base64Encode(urn);
            if (String.IsNullOrEmpty(viewableId))//vista 3D               
                                                 //respuesta = string.Format("file:///HTML/Viewer.html?URN={0}&Token={1}", Base64Encode(urn), txtAccessToken.Text);
                respuesta = string.Format(@"C:\HTML1\Viewer.html?URN={0}&Token={1}", Base64Encode(urn), txtAccessToken.Text);

            else
                // respuesta = string.Format("file:///HTML/Viewer.html?URN={0}&Token={1}&ViewableId={2}", Base64Encode(urn), txtAccessToken.Text, viewableId);
                //respuesta = string.Format(@"C:\HTML1\Viewer.html?URN={0}&Token={1}&ViewableId={2}", Base64Encode(urn), txtAccessToken.Text, viewableId);
                respuesta = string.Format(@"C:\HTML1\Viewer.html?URN={0}&Token={1}&ViewableId={2}", Base64Encode(urn), txtAccessToken.Text, viewableId);
            return respuesta;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            //AccessToken ELEMENTO= new AccessToken();
            //txtAccessToken.Text= ELEMENTO.Token(txtClientId.Text, txtClientSecret.Text).access_token;

        }
    }
}
