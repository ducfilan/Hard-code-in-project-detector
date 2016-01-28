using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using RF.Properties;
using System.Text.RegularExpressions;
using System.Xml;
using RF.Parser;
using System.CodeDom.Compiler;

namespace RF
{
    public partial class Form1 : Form
    {
        List<string> _fileList = new List<string>();

        public Form1()
        {
            InitializeComponent();
            Load += Form1_Load;
            dataGridView1.SelectionChanged += dataGridView1_SelectionChanged;
        }

        void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0) return;

            var boundItem = dataGridView1.SelectedRows[0].DataBoundItem as Item;

            Clear();

            originalTxt.Text = boundItem.ExtractedValue;
            if (!string.IsNullOrEmpty(boundItem.Suggestion))
            {
                suggetionTxt.Text = boundItem.Suggestion;
            }

        }

        void Form1_Load(object sender, EventArgs e)
        {

            ImageList iconList = new ImageList();
            iconList.Images.Add(Resources.sln);
            iconList.Images.Add(Resources.prj);
            iconList.Images.Add(Resources.folder);
            iconList.Images.Add(Resources.aspx);
            iconList.Images.Add(Resources.ascx);
            iconList.Images.Add(Resources.vb);
            iconList.Images.Add(Resources.js);
            iconList.Images.Add(Resources.tick);
            iconList.Images.Add(Resources.ukn);

            treeView1.AfterSelect += treeView1_AfterSelect;
            treeView1.ImageList = iconList;
        }

        List<Item> selectedItemList;

        void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (treeView1.SelectedNode == null) return;

            var path = treeView1.SelectedNode.Tag as string;

            Task t = Task.Factory.StartNew(() =>
            {
                try
                {
                    // do your processing here - remember to call Invoke or BeginInvoke if
                    // calling a UI object.
                    Invoke((MethodInvoker)delegate
                    {
                        treeView1.Enabled = false;
                        dataGridView1.Enabled = false;
                        toolStripProgressBar1.Style = ProgressBarStyle.Marquee;
                        toolStripProgressBar1.MarqueeAnimationSpeed = 30;
                        toolStripLabel1.Text = "Parsing file. Please wait...";
                    });

                    List<string> oFiles = null;

                    var selectedNode = GetSelectedNode();
                    if (selectedNode != null)
                    {
                        string f;
                        if ((f = _fileList.FirstOrDefault(p => p == selectedNode.Tag)) != null)
                        {
                            oFiles = new List<string>();
                            oFiles.Add(f);
                        }
                        else
                        {
                            oFiles = _fileList.Where(p => p.Contains(selectedNode.Tag + string.Empty)).ToList();
                        }
                    }

                    if (oFiles == null || oFiles.Count == 0)
                    {
                        return;
                    }

                    selectedItemList = new List<Item>();

                    foreach (var f in oFiles)
                    {
                        var pd = ParseFile(f);
                        if (pd != null && pd.Count > 0)
                        {
                            selectedItemList.AddRange(pd);
                        }
                    }

                    Invoke((MethodInvoker)delegate
                    {
                        dataGridView1.DataSource = selectedItemList;


                        Clear();


                        if (selectedItemList != null && selectedItemList.Count > 0)
                        {
                            originalTxt.Text = selectedItemList[0].ExtractedValue;
                            if (!string.IsNullOrEmpty(selectedItemList[0].Suggestion))
                            {
                                suggetionTxt.Text = selectedItemList[0].Suggestion;
                            }

                        }
                    });
                }
                catch (Exception)
                {

                }

            });

            // ReSharper disable ImplicitlyCapturedClosure
            t.ContinueWith(success =>
                // ReSharper restore ImplicitlyCapturedClosure
                // Export to excel
                Invoke((MethodInvoker)delegate
                {
                    toolStripProgressBar1.MarqueeAnimationSpeed = 0;
                    toolStripProgressBar1.Style = ProgressBarStyle.Blocks;
                    toolStripLabel1.Text = "Parse successfully.";
                    treeView1.Enabled = true;
                    dataGridView1.Enabled = true;
                }), TaskContinuationOptions.NotOnFaulted);
            t.ContinueWith(fail =>
            {
                //log the exception i.e.: Fail.Exception.InnerException);
                Invoke((MethodInvoker)delegate
                {
                    toolStripProgressBar1.MarqueeAnimationSpeed = 0;
                    toolStripProgressBar1.Style = ProgressBarStyle.Blocks;
                    toolStripLabel1.Text = "Error";
                    treeView1.Enabled = true;
                    dataGridView1.Enabled = true;
                });
            }, TaskContinuationOptions.OnlyOnFaulted);


        }

        private string _basedPath;
        TreeNodeSimulation _solutionNode;

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = Settings.Default["InitialDirectory"].ToString();

            openFileDialog1.Filter = "Solution File (*.sln)|*.sln";
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                Task t = Task.Factory.StartNew(() =>
                {
                    Invoke((MethodInvoker)delegate
                    {
                        toolStripProgressBar1.Style = ProgressBarStyle.Marquee;
                        toolStripProgressBar1.MarqueeAnimationSpeed = 30;
                        toolStripLabel1.Text = "Loading...";

                        if (treeView1.Nodes.Count > 0)
                        {
                            for (int i = treeView1.Nodes.Count - 1; i >= 0; i--)
                            {
                                treeView1.Nodes[i].Remove();
                            }
                        }

                        _fileList = new List<string>();
                    });

                    _basedPath = Path.GetDirectoryName(openFileDialog1.FileName);

                    Settings.Default["InitialDirectory"] = _basedPath;
                    Settings.Default.Save();

                    _solutionNode = new TreeNodeSimulation(Path.GetFileName(openFileDialog1.FileName), openFileDialog1.FileName, 0);

                    var projectCount = 0;
                    try
                    {
                        Regex regex = new Regex("Project.* = ([^,]*),([^,]*)");
                        System.IO.StreamReader file = new System.IO.StreamReader(openFileDialog1.FileName);

                        string line;
                        while ((line = file.ReadLine()) != null)
                        {
                            var match = regex.Match(line);
                            if (match.Success)
                            {
                                var path = Path.Combine(_basedPath, match.Value
                                                                    .Substring(match.Value.LastIndexOf(',') + 1)
                                                                    .Replace("\"", "")
                                                                    .Trim());
                                projectCount++;
                                var projectBasedPath = Path.GetDirectoryName(path);

                                var projNode = new TreeNodeSimulation(Path.GetFileName(path), path, 1);
                                _solutionNode.Nodes.Add(projNode);
                                var doc = new XmlDocument();
                                doc.Load(path);

                                var nsmgr = new XmlNamespaceManager(doc.NameTable);
                                nsmgr.AddNamespace("ms", "http://schemas.microsoft.com/developer/msbuild/2003");
                                var includedFiles = doc.SelectNodes(@"//ms:Compile[@Include] | //ms:Content[@Include]", nsmgr);

                                foreach (var f in includedFiles)
                                {
                                    var fnode = f as XmlNode;
                                    var n = fnode.Attributes.GetNamedItem("Include");
                                    var temp = n.Value;

                                    var segments = temp.Split(new[] { "\\" }, StringSplitOptions.RemoveEmptyEntries);
                                    if (segments[segments.Length - 1].Length <= 12 ||
                                        (segments[segments.Length - 1].Substring(segments[segments.Length - 1].Length - 12, 12).ToLower() != ".designer.vb" &&
                                        Path.GetExtension(segments[segments.Length - 1]).ToLower() != ".resx" &&
                                        Path.GetExtension(segments[segments.Length - 1]).ToLower() != ".dll" &&
                                        segments[segments.Length - 1].ToLower() != "assemblyinfo.vb"))
                                    {
                                        if (segments.Length == 1)
                                        {

                                            var fNode = new TreeNodeSimulation(segments[0], Path.Combine(projectBasedPath, temp), GetImageIndex(segments[0]));
                                            _fileList.Add(fNode.Tag + string.Empty);
                                            projNode.Nodes.Add(fNode);
                                        }
                                        else
                                        {
                                            var parN = projNode;

                                            for (var i = 0; i < segments.Length - 1; i++)
                                            {
                                                var curN = parN.Nodes.FirstOrDefault(p => p.Name == segments[i]);

                                                if (curN == null)
                                                {

                                                    var currPath = string.Empty;
                                                    for (var j = 0; j <= i; j++)
                                                    {
                                                        currPath += segments[j] + "\\";
                                                    }

                                                    curN = new TreeNodeSimulation(segments[i], Path.Combine(projectBasedPath, currPath), 2);
                                                    parN.Nodes.Add(curN);
                                                }

                                                parN = curN;
                                            }

                                            var fileName = segments[segments.Length - 1];
                                            var leafN = new TreeNodeSimulation(fileName, Path.Combine(projectBasedPath, temp), GetImageIndex(fileName));
                                            _fileList.Add(leafN.Tag + string.Empty);
                                            parN.Nodes.Add(leafN);
                                        }
                                    }
                                }
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        Invoke((MethodInvoker)delegate
                        {
                            toolStripProgressBar1.Style = ProgressBarStyle.Blocks;
                            toolStripProgressBar1.MarqueeAnimationSpeed = 0;
                            toolStripLabel1.Text = "Error: Could not read file from disk.";
                        });
                    }

                    SortNodes(_solutionNode);

                    Invoke((MethodInvoker)delegate
                    {
                        BindTree(_solutionNode, treeView1.Nodes);

                        toolStripProgressBar1.Style = ProgressBarStyle.Blocks;
                        toolStripProgressBar1.MarqueeAnimationSpeed = 0;
                        toolStripLabel1.Text = "Loaded successfully " + _fileList.Count + " file(s) in " + projectCount + " project(s).";
                    });
                });

            }
        }

        private void SortNodes(TreeNodeSimulation tns)
        {
            if (tns.Nodes.Count == 0)
            {
                return;
            }

            tns.Nodes.Sort((p1, p2) => string.Compare(p1.Name, p2.Name, true));

            foreach (var t in tns.Nodes)
            {
                SortNodes(t);
            }
        }

        private void BindTree(TreeNodeSimulation visual, TreeNodeCollection nodes)
        {
            var n = new TreeNode(visual.Name);
            n.Name = visual.Name;
            n.Tag = visual.Tag;
            n.SelectedImageIndex = 7;
            n.ImageIndex = visual.ImageIndex;

            nodes.Add(n);

            if (visual.Nodes.Count == 0)
                return;

            foreach (var child in visual.Nodes)
            {
                BindTree(child, n.Nodes);
            }
        }

        private int GetImageIndex(string fileName)
        {
            var index = fileName.LastIndexOf(".");
            var extension = fileName.Substring(index, fileName.Length - index);

            switch (extension.ToLower())
            {
                case ".cs":
                    return 5;

                case ".js":
                    return 6;

                case ".aspx":
                    return 3;

                case ".ascx":
                    return 4;

                default:
                    return 8;
            }
        }

        AspParser _aspParser = new AspParser();
        DotNetParser _dotNetParser = new DotNetParser();
        JsParser _jsParser = new JsParser();

        private List<Item> ParseFile(string path)
        {
            var extension = Path.GetExtension(path);

            IParser parser = null;

            switch (extension.ToLower())
            {
                case ".ascx":
                case ".aspx":
                case ".master":
                    parser = _aspParser;
                    break;
                
                case ".cs":
                case ".vb":
                    parser = _dotNetParser;
                    break;

                //optional ?
                case ".js":
                    parser = _jsParser;
                    break;
            }


            if (parser == null)
            {
                return null;
            }

            var boundList = parser.GetHardCodeList(path);

            boundList.Sort((p1, p2) => (p1.Line- p2.Line));
            return boundList;
        }

        private void Clear()
        {
            originalTxt.Text = string.Empty;
            suggetionTxt.Text = string.Empty;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Task t = Task.Factory.StartNew(() =>
            {
                try
                {
                    // do your processing here - remember to call Invoke or BeginInvoke if
                    // calling a UI object.
                    Invoke((MethodInvoker)delegate
                    {
                        toolStripProgressBar1.Style = ProgressBarStyle.Marquee;
                        toolStripProgressBar1.MarqueeAnimationSpeed = 30;
                        toolStripLabel1.Text = "Generating report..";
                    });


                    using (var tempFiles = new TempFileCollection())
                    {
                        string file = tempFiles.AddExtension("xlsx");
                        // do something with the file here 
                        ExcelHelper.SaveReport(selectedItemList, file);
                    }

                }
                catch (Exception)
                {

                }

            });

            // ReSharper disable ImplicitlyCapturedClosure
            t.ContinueWith(success =>
                // ReSharper restore ImplicitlyCapturedClosure

                Invoke((MethodInvoker)delegate
                {
                    toolStripProgressBar1.MarqueeAnimationSpeed = 0;
                    toolStripProgressBar1.Style = ProgressBarStyle.Blocks;
                    toolStripLabel1.Text = "Exported successfully.";
                }), TaskContinuationOptions.NotOnFaulted);
            t.ContinueWith(fail =>
            {
                //log the exception i.e.: Fail.Exception.InnerException);
                Invoke((MethodInvoker)delegate
                {
                    toolStripProgressBar1.MarqueeAnimationSpeed = 0;
                    toolStripProgressBar1.Style = ProgressBarStyle.Blocks;
                    toolStripLabel1.Text = "Error";
                });
            }, TaskContinuationOptions.OnlyOnFaulted);

        }

        // This delegate enables asynchronous calls for setting
        // the text property on a TextBox control.
        delegate TreeNode SetTextCallback();

        private TreeNode GetSelectedNode()
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.treeView1.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(GetSelectedNode);
                return this.Invoke(d) as TreeNode;
            }
            else
            {
                return treeView1.SelectedNode;
            }
        }


        private void button4_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode != null)
                System.Diagnostics.Process.Start(treeView1.SelectedNode.Tag + string.Empty);
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                var item = dataGridView1.SelectedRows[0].DataBoundItem as Item;

                item.Suggestion = suggetionTxt.Text;

                dataGridView1.Refresh();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            foreach (var r in dataGridView1.Rows)
            {
                var row = r as DataGridViewRow;
                if (Convert.ToBoolean(row.Cells[0].FormattedValue) == true)
                {
                    SaveItemToFile(row.DataBoundItem as Item);
                }
            }
        }

        private void SaveItemToFile(Item item)
        {
            if (item == null)
                return;

            var content = System.IO.File.ReadAllText(item.FilePath);

            var lines = content.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            lines[item.Line - 1] = lines[item.Line - 1].Replace(item.ExtractedValue, item.Suggestion);

            var newContent = string.Join(Environment.NewLine, lines);

            try
            {
                using (StreamWriter outfile = new StreamWriter(item.FilePath, false))
                {
                    outfile.Write(newContent);
                }

                MessageBox.Show("Updated");
            }
            catch
            {
                MessageBox.Show("Error !!");
            }
        }
    }

}
