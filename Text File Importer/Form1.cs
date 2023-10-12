using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using OneNote = Microsoft.Office.Interop.OneNote;
using System.Text.RegularExpressions;
using HtmlAgilityPack;


namespace Text_File_Importer
{
    public partial class FormTextImporter : Form
    {
        string sectionID; //section where the new pages will be created
        XmlDocument xmlDoc = new XmlDocument();
        static String strNamespace = "http://schemas.microsoft.com/office/onenote/2007/onenote";
        private string m_xmlNewOutlineContent =
            "<one:Meta name=\u0022{2}\u0022 content=\u0022{1}\u0022/>" +
            "<one:OEChildren><one:HTMLBlock><one:Data><![CDATA[{0}]]></one:Data></one:HTMLBlock></one:OEChildren>";
        private string m_xmlNewOutline =
           "<?xml version=\u00221.0\u0022?>" +
           "<one:Page xmlns:one=\u0022{2}\u0022 ID=\u0022{1}\u0022>" +
           "<one:Title><one:OE><one:T><![CDATA[{3}]]></one:T></one:OE></one:Title>" +
           "<one:Outline>{0}</one:Outline></one:Page>";

        string[] notebookID, pathToNotebook;
        private string m_outlineIDMetaName = "OneNote Text File Importer";
        bool notebookSelected = false;
        string pathToTextFiles, newSectionName;
        System.IO.StreamReader myFile;
        System.IO.StreamWriter logFile = null;
        OneNote.Application onApp = new OneNote.Application();

        public FormTextImporter()
        {
            InitializeComponent();
            notebookID = new string[5000];
            pathToNotebook = new string[5000];
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            progressBar1.Visible = false;
            textInstructions.Text = "Browse to the folder which holds the text files you want to import.";
            textNotebookPicker.Text = "Select a notebook to use for importing.  If you do not, your files will be imported to Unfiled Notes.";
            try
            {
                logFile = new StreamWriter(System.Environment.CurrentDirectory + "\\OneNoteTextFile_Importer_error_log.txt");
                UpdateTree();
            }
            catch { }
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                pathToTextFiles = folderBrowserDialog1.SelectedPath;
                textBox1.SelectAll();
                textBox1.SelectedText = pathToTextFiles;
            }

        }

        private void buttonImport_Click(object sender, EventArgs e)
        {
            string strPath = "", result = "Unfiled Notes";
            try
            {
                if (notebookSelected == true)
                {
                    onApp.NavigateTo(notebookID[treeView1.SelectedNode.Index], System.String.Empty, false);
                    strPath = pathToNotebook[treeView1.SelectedNode.Index];
                    result = treeView1.SelectedNode.Text;
                }
                else //put new pages in Unfiled Notes 
                {
                    onApp.GetSpecialLocation(Microsoft.Office.Interop.OneNote.SpecialLocation.slUnfiledNotesSection, out strPath);
                    onApp.OpenHierarchy(strPath, "", out sectionID, OneNote.CreateFileType.cftNone);
                    onApp.NavigateTo(sectionID, System.String.Empty, false);
                }
                pathToTextFiles = textBox1.Text;

                //Willem added SearchOption.AllDirectories.
                string[] fileName = Directory.GetFiles(pathToTextFiles, "*.txt", SearchOption.AllDirectories);

                if (fileName.Length == 0)
                {
                    MessageBox.Show("There are no text files at to import at " + pathToTextFiles, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                DirectoryInfo dirInfo = Directory.GetParent(fileName[0]);
                string dirName = dirInfo.Name;
                //need to clean the dirName of illegal chars
                string d1 = dirName.Replace(':', '\\');
                dirName = d1.TrimEnd('\\');
                int fileCount = fileName.Length;
                progressBar1.Visible = true;
                progressBar1.Maximum = fileCount;

                strPath += "\\" + dirName + ".one";
                if (notebookSelected == true)
                {
                    onApp.OpenHierarchy(strPath, null, out sectionID, OneNote.CreateFileType.cftSection);
                    onApp.NavigateTo(sectionID, "", false);
                }
                int i;
                String tempLine;
                for (i = 0; i < fileCount; i++)
                {
                    //iterate through files and import
                    myFile = new System.IO.StreamReader(fileName[i], System.Text.Encoding.Default);
                    int count = 0;

                    StringBuilder sb = new StringBuilder();
                    while ((tempLine = myFile.ReadLine()) != null)
                    {
                        Encoding targetEncoding = myFile.CurrentEncoding;
                        String lineToAppend = "";
                        foreach (byte b in targetEncoding.GetBytes(tempLine))
                            lineToAppend += Convert.ToChar(b);
                        sb.AppendLine(lineToAppend);
                        //sb.AppendLine(tempLine);
                        count++;
                    }
                    //get the name of the page

                    string[] fn = fileName[i].Split('\\');
                    string newPageName = fn[fn.Length - 1].Substring(0, fn[fn.Length - 1].Length - 4);//assumes .txt at end
                    myFile.Close();

                    #region CREATEPAGE
                    //need to convert /r/n to <br>  since we need html tags for OneNote
                    string pText = sb.Replace("\r\n", "<BR>").ToString();
                    pText = "<html><body>" + pText + "</body></html>";

                    XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                    nsmgr.AddNamespace("one", strNamespace);
                    //Create a page in the given sectionID
                    string strNewPageID;
                    onApp.CreateNewPage(sectionID, out strNewPageID, OneNote.NewPageStyle.npsBlankPageWithTitle);
                    int outlineID = new System.Random().Next();
                    //take the xml strings and add in the data unique to this book being imported
                    string outlineContent = string.Format(m_xmlNewOutlineContent, pText, outlineID, m_outlineIDMetaName);
                    string xml = string.Format(m_xmlNewOutline, outlineContent, strNewPageID, strNamespace, newPageName);
                    onApp.UpdatePageContent(xml, DateTime.MinValue);
                    #endregion 
                    progressBar1.Increment(1);
                }
                MessageBox.Show("Successfully imported " + i.ToString() +
                                " files to " + result + ".", "Done!",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                progressBar1.Value = 0;
                buttonBrowse.Focus();
            }
            catch (Exception clickException)
            {
                logFile.WriteLine(clickException.ToString());
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            notebookSelected = true;
            TreeNode temp = treeView1.SelectedNode;
            newSectionName = temp.Text;
            if (newSectionName.Equals("Notebooks"))
            {
                newSectionName = "";
                notebookSelected = false;
            }
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            string strPath = "", result = "Unfiled Notes";
            onApp.NavigateTo(notebookID[treeView1.SelectedNode.Index], System.String.Empty, false);
            strPath = pathToNotebook[treeView1.SelectedNode.Index];
            result = treeView1.SelectedNode.Text;

            // このNotebook内のすべてのページIDを取得する処理を行う
            string xmlHierarchy;
            string notebookId = notebookID[treeView1.SelectedNode.Index];
            onApp.GetHierarchy(notebookId, OneNote.HierarchyScope.hsPages, out xmlHierarchy);
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlHierarchy);

            XmlNamespaceManager nsMgr = new XmlNamespaceManager(doc.NameTable);
            nsMgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");

            // 保存のルートディレクトリ
            string rootDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Notebookのフォルダを作成
            string notebookDirectory = Path.Combine(rootDirectory, result); // resultはNotebookの名前
            Directory.CreateDirectory(notebookDirectory);

            var htmlDoc = new HtmlAgilityPack.HtmlDocument();


            // 新しく追加するコード
            XmlNodeList sectionNodes = doc.SelectNodes("//one:Section", nsMgr);
            foreach (XmlNode sectionNode in sectionNodes)
            {
                string sectionName = sectionNode.Attributes["name"].Value;
                Console.WriteLine("Processing section: " + sectionName);  // 何らかの処理

                // セクションのフォルダを作成
                string sectionDirectory = Path.Combine(notebookDirectory, sectionName);
                Directory.CreateDirectory(sectionDirectory);

                XmlNodeList pageNodes = sectionNode.SelectNodes("one:Page", nsMgr);
                foreach (XmlNode page in pageNodes)
                {
                    string pageName = page.Attributes["name"].Value;
                    Console.WriteLine("  Processing page: " + pageName);  // 何らかの処理

                    string pageId = page.Attributes["ID"].Value;

                    // ページの内容を取得
                    string pageContentXml;
                    onApp.GetPageContent(pageId, out pageContentXml, OneNote.PageInfo.piAll);
                    XmlDocument pageDoc = new XmlDocument();
                    pageDoc.LoadXml(pageContentXml);

                    XmlNamespaceManager nsMgrPage = new XmlNamespaceManager(pageDoc.NameTable);
                    nsMgrPage.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");

                    // ページのタイトルを取得（2番目のone:Tを選ぶ）
                    XmlNodeList titleNodes = pageDoc.SelectNodes("//one:Title//one:T", nsMgrPage);
                    XmlNode titleNode = titleNodes.Count > 1 ? titleNodes[1] : (titleNodes.Count > 0 ? titleNodes[0] : null);  // タイトルが2つ以上なら2番目、それ以外は1番目かnull
                    string titleText = titleNode != null ? titleNode.InnerText : "Untitled";  // タイトルがない場合は"Untitled"

                    // タイトルからHTML/XMLタグを削除
                    titleText = Regex.Replace(titleText, "\r\n", String.Empty);
                    titleText = Regex.Replace(titleText, "<.*?>", String.Empty);

                    // ページの本文を取得
                    StringBuilder pageContent = new StringBuilder();
                    XmlNodeList outlineNodes = pageDoc.SelectNodes("//one:Outline//one:T", nsMgrPage);
                    foreach (XmlNode outlineNode in outlineNodes)
                    {
                        if (outlineNode.InnerText != null)
                        {
                            string outlineText = outlineNode.InnerText;
                            htmlDoc.LoadHtml(outlineText);
                            var outlineText2 = HtmlEntity.DeEntitize(htmlDoc.DocumentNode.InnerText);
                            pageContent.AppendLine(outlineText2);
                        }
                    }

                    // 保存先のテキストファイル（この行を修正）
                    string filePath = Path.Combine(sectionDirectory, titleText + ".txt");

                    // テキストファイルに書き出し
                    File.WriteAllText(filePath, pageContent.ToString());

                }
            }

            MessageBox.Show("エクスポート完了！対象のNotebookは " + result + "だぞ。", "完成！", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        /// <summary>
        /// shows the existing notebooks & unfiled notes in the treeview
        /// </summary>
        private void UpdateTree()
        {
            try
            {
                System.Xml.XmlDocument dom = new System.Xml.XmlDocument();
                OneNote.Application onApp = new OneNote.Application();

                // Get the hierarchy for all notebooks to Notebook level
                string strHierarchy;
                onApp.GetHierarchy(System.String.Empty, OneNote.HierarchyScope.hsNotebooks, out strHierarchy);
                dom.LoadXml(strHierarchy);

                treeView1.Nodes.Clear();
                treeView1.Nodes.Add(new TreeNode(dom.DocumentElement.LocalName));
                TreeNode tNode = new TreeNode();
                tNode = treeView1.Nodes[0];
                AddNode(dom.DocumentElement, tNode);
            }
            catch (Exception ex)
            {
                if (logFile != null)
                    logFile.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// adds nodes to the treeview control
        /// </summary>
        /// <param name="inXmlNode"></param>
        /// <param name="inTreeNode"></param>
        private void AddNode(System.Xml.XmlNode inXmlNode, TreeNode inTreeNode)
        {
            try
            {
                System.Xml.XmlNode xNode;
                TreeNode tNode;
                System.Xml.XmlNodeList nodeList;
                int i;

                // Loop through the XML nodes until the notebooks are reached.
                // Add the nodes to the TreeView during the looping process.
                System.Xml.XmlAttributeCollection a;
                XmlNode nm;
                if (inXmlNode.HasChildNodes)
                {
                    nodeList = inXmlNode.ChildNodes;
                    for (i = 0; i <= nodeList.Count - 1; i++)
                    {
                        xNode = inXmlNode.ChildNodes[i];
                        inTreeNode.Nodes.Add(new TreeNode(xNode.Name));
                        tNode = inTreeNode.Nodes[i];

                        AddNode(xNode, tNode);
                        a = xNode.Attributes;
                        if (tNode.Text.Contains("Unfiled") == false)
                        {
                            tNode.Text = a.GetNamedItem("name").Value;
                            nm = a.GetNamedItem("ID");
                            notebookID[i] = nm.Value;
                            nm = a.GetNamedItem("path");
                            pathToNotebook[i] = nm.Value;

                        }
                        else
                        {
                            tNode.Text = "Unfiled Notes";
                            nm = a.GetNamedItem("ID");
                            notebookID[i] = nm.Value;
                            //don't need the path to the Unfiled notes notebook: use
                            //getspeciallocation when creating chapters
                        }
                    }
                }
                else
                {
                    inTreeNode.Text = (inXmlNode.OuterXml).Trim();
                }
            }
            catch (Exception ex)
            {
                if (logFile != null)
                    logFile.WriteLine(ex.ToString());
            }

        }
    }
}
