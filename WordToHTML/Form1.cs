using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Text.RegularExpressions;

namespace WordToHTML
{
    public partial class Form1 : Form
    {
        private static string docPath = "";
        private static string fileType = "";
        private static string ftprootpath = "";
        private static string[] ftpDirs;
        private static string[] uploadImgs;
        private static string webAdd = @"http://www.madefafa.top";

        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (docPath != "")
            {
                string[] pathFormat = docPath.Split('\\');
                fileType = (pathFormat[pathFormat.Length - 1].Split('.'))[1];
                WordToHtmlFile(docPath);
                groupBox2.Enabled = true;
                MessageBox.Show("文档处理完毕!");

            }

        }

        private static void WordToHtmlFile(string WordFilePath)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application newApp = new Microsoft.Office.Interop.Word.Application();
                //指定原文件和目标文件
                object Source = WordFilePath;
                
                string SaveHtmlPath = WordFilePath.Substring(0, WordFilePath.Length - fileType.Length) + "html";
                object Target = SaveHtmlPath;

                // 缺省参数  
                object Unknown = Type.Missing;

                //为了保险,只读方式打开
                object readOnly = true;

                // 打开doc文件
                Microsoft.Office.Interop.Word.Document doc = newApp.Documents.Open(ref Source, ref Unknown,
                     ref readOnly, ref Unknown, ref Unknown,
                     ref Unknown, ref Unknown, ref Unknown,
                     ref Unknown, ref Unknown, ref Unknown,
                     ref Unknown, ref Unknown, ref Unknown,
                     ref Unknown, ref Unknown);

                // 指定另存为格式(rtf)
                object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatFilteredHTML;
                // 转换格式
                doc.SaveAs(ref Target, ref format,
                        ref Unknown, ref Unknown, ref Unknown,
                        ref Unknown, ref Unknown, ref Unknown,
                        ref Unknown, ref Unknown, ref Unknown,
                        ref Unknown, ref Unknown, ref Unknown,
                        ref Unknown, ref Unknown);

                // 关闭文档和Word程序
                doc.Close(ref Unknown, ref Unknown, ref Unknown);
                newApp.Quit(ref Unknown, ref Unknown, ref Unknown);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
            }

         }

        private void btnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory =@"C:\Users\Administrator\Desktop";
            ofd.Filter = ".docx(Word文档)|*.docx|.doc(Word 93-97)|*.doc";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                docPath = ofd.FileName;
                textBox1.Text = docPath;
                MessageBox.Show("所需文档已经导入完毕！");
                btnConvert.Enabled = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            btnConvert.Enabled = false;
           // groupBox2.Enabled = false;
        }



        private void btnUpload_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(txtbFtpPath.Text) || String.IsNullOrEmpty(txtbUsrId.Text) || String.IsNullOrEmpty(txtbPWD.Text))
            {
                MessageBox.Show("远程服务器登录信息不完整!", "赵毅工作室2017", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                return;
            }

            ftprootpath = txtbFtpPath.Text;
            string uploadPath = "";
            
            for (int i = 0; i < docPath.Split('\\').Length - 1; i++)
            {
                uploadPath += docPath.Split('\\')[i]+"\\";
            }
          string[] filename= docPath.Split('\\')[docPath.Split('\\').Length - 1].Split('.');
            uploadPath += filename[0] + ".files";

            FtpHelper fhp = new FtpHelper();
            ftpDirs = fhp.GetFilePath(txtbUsrId.Text, txtbPWD.Text, txtbFtpPath.Text);
            if (fhp.DirectoryExist(filename[0])) //首先判断是否存在该子文件夹(包含图片的),如果已经存在报错返回.
            {
                MessageBox.Show("业已存在名为" + filename[0] + "的文件夹\n请重新命名!");
                return;
            }
            else
            {

                //如果没有存在,创建该文件夹
                fhp.MakeDir(filename[0]);
                uploadImgs = Directory.GetFiles(uploadPath);

                //然后循环把文件夹内图片上传纸
                foreach (string str in uploadImgs)
                {
                    fhp.Upload(txtbUsrId.Text, txtbPWD.Text, str, txtbFtpPath.Text + "/" + filename[0]);
                }
                lblTransInfo .Text= "传输数量:" + uploadImgs.Length;
                MessageBox.Show("所有文件上传完毕!");
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            

            string[] webpathCont= ftprootpath.Split('/');
            for (int i = 4; i < webpathCont.Length; i++)
            {
                webAdd += "/" + webpathCont[i];
            }

            string html = docPath.Split('.')[0] + ".html";
           string content= File.ReadAllText(html,Encoding.Default);
            //第一步把生成的HTML<body></body>前后的东西全搞掉
            content = Regex.Match(content, @"<body[\s\S]*</body>").ToString();
            //第二步把文本框中的正则表达式运算,替换照片路径
            string regStr = rtbRegExp.Text;
          MatchCollection output=  Regex.Matches(content, regStr);
            for (int i = 0; i < output.Count; i++)
            {
                string[] repStr = uploadImgs[i].Split('\\');
                string newStr = webAdd + "/" + repStr[repStr.Length - 2].Split('.')[0] + "/" + repStr[repStr.Length - 1];  
                content = content.Replace(output[i].ToString(), richTextBox2.Text.Insert(richTextBox2.Text.IndexOf('"')+1,newStr));
                
            }
            //第三步根据需要给body主体加表格
            //1.删除<body...>
            content = Regex.Replace(content, @"<body[^<]*>", "");
            //2.删除</body>
            content = Regex.Replace(content, @"</body>", "");
            //依据textbox是否为空给其加表格
            if (!String.IsNullOrEmpty(txtbWidth.Text))
            {
                //如果不为在外层套上表格
                string upCode = "<table width=\""+ txtbWidth.Text + "\" border=\"0\" align=\"center\"> \n <tr> \n <td>\n";
                string downCode = "\n </td> \n </tr> \n </table> ";
                content = upCode + content + downCode;
            }


            //string styleStr = String.Format(@"<style>[\s\S]*</style>");
            //string strReplace = @"<style> body "+'{'+ "margin: 0 auto; width: "+txtbWidth.Text +"px;"+'}'+" </style>";
            // Regex.Replace(content, styleStr, strReplace);
           // content=Regex.Replace(content, styleStr, strReplace);
            File.WriteAllText(html, content,Encoding.Default);


            MessageBox.Show("文档已保存!", "赵毅工作室2017", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnlogin_Click(object sender, EventArgs e)
        {
            if (txtbUser.Text.Trim() == "user" && txtbPass.Text.Trim() == "123456")
            {
                MessageBox.Show("登录成功!", "赵毅工作室2017", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtbFtpPath.Text= @"此处填写FTP地址";
                txtbUsrId.Text = @"此处填写FTP登录用户名";
                txtbPWD.Text = @"此处填写FTP密码";
            }
        }
    }
    }

