using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace TestDirectory
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox1.Text = @"D:\FiscaAEAlpha";        }

        string _Path = "";
        private void button1_Click(object sender, EventArgs e)
        {
            _Path = textBox1.Text;

            if (Directory.Exists(_Path))
            {
                try
                {
                    DirectoryInfo dis = new DirectoryInfo(_Path);  
                    treeView1.Nodes.Clear();
                    var stack = new Stack<TreeNode>();
                    TreeNode node = new TreeNode(dis.Name);
                    node.Tag=dis;
                    stack.Push(node);

                    while (stack.Count > 0)
                    {
                        TreeNode curNote = stack.Pop();
                        DirectoryInfo difo = curNote.Tag as DirectoryInfo;
                        foreach (DirectoryInfo dif in difo.GetDirectories())
                        {
                            TreeNode tn = new TreeNode(dif.Name);
                            tn.Tag = dif;
                            curNote.Nodes.Add(tn);
                            stack.Push(tn);
                        }

                        foreach (FileInfo fi in difo.GetFiles())
                            curNote.Nodes.Add(new TreeNode(fi.Name));
                    }
                    treeView1.Nodes.Add(node);
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("目錄不存在");
            }

        }
    }
}
