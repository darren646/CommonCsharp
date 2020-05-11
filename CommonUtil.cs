using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.DirectoryServices.AccountManagement;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;

namespace CommonUtil
{

    public class CSVFileHelper
    {
        /// <summary>
        /// 将DataTable中数据写入到CSV文件中
        /// </summary>
        /// <param name="dt">提供保存数据的DataTable</param>
        /// <param name="fileName">CSV的文件路径</param>
        public static void WriteCSV(DataTable dt, string fullPath)
        {
            FileInfo fi = new FileInfo(fullPath);
            if (!fi.Directory.Exists)
            {
                fi.Directory.Create();
            }
            FileStream fs = new FileStream(fullPath, System.IO.FileMode.Create, System.IO.FileAccess.Write);
            //StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.Default);
            StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8);
            string data = "";
            //写出列名称
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                data += dt.Columns[i].ColumnName.ToString();
                if (i < dt.Columns.Count - 1)
                {
                    data += ",";
                }
            }
            sw.WriteLine(data);
            //写出各行数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                data = "";
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    string str = dt.Rows[i][j].ToString();
                    str = str.Replace("\"", "\"\"");//替换英文冒号 英文冒号需要换成两个冒号
                    if (str.Contains(',') || str.Contains('"')
                        || str.Contains('\r') || str.Contains('\n')) //含逗号 冒号 换行符的需要放到引号中
                    {
                        str = string.Format("\"{0}\"", str);
                    }

                    data += str;
                    if (j < dt.Columns.Count - 1)
                    {
                        data += ",";
                    }
                }
                sw.WriteLine(data);
            }
            sw.Close();
            fs.Close();

            Console.WriteLine("Write csv file successfully!");

            //DialogResult result = MessageBox.Show("CSV文件保存成功！");
            //if (result == DialogResult.OK)
            //{
            //    System.Diagnostics.Process.Start("explorer.exe");
            //}
        }

        /// <summary>
        /// 将CSV文件的数据读取到DataTable中
        /// </summary>
        /// <param name="fileName">CSV文件路径</param>
        /// <returns>返回读取了CSV数据的DataTable</returns>
        public static DataTable ReadCSV(string filePath)
        {
            //Encoding encoding = Encoding.GetEncoding(filePath); //Encoding.ASCII;//
            DataTable dt = new DataTable();
            FileStream fs = new FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read,FileShare.ReadWrite);

            StreamReader sr = new StreamReader(fs, Encoding.ASCII);
            //StreamReader sr = new StreamReader(fs, encoding);
            //string fileContent = sr.ReadToEnd();
            //encoding = sr.CurrentEncoding;
            //记录每次读取的一行记录
            string strLine = "";
            //记录每行记录中的各字段内容
            string[] aryLine = null;
            string[] tableHead = null;
            //标示列数
            int columnCount = 0;
            //标示是否是读取的第一行
            bool IsFirst = true;
            //逐行读取CSV中的数据
            while ((strLine = sr.ReadLine()) != null)
            {
                //strLine = Common.ConvertStringUTF8(strLine, encoding);
                //strLine = Common.ConvertStringUTF8(strLine);

                if (IsFirst == true)
                {
                    tableHead = strLine.Split(',');
                    IsFirst = false;
                    columnCount = tableHead.Length;
                    //创建列
                    for (int i = 0; i < columnCount; i++)
                    {
                        DataColumn dc = new DataColumn(tableHead[i]);
                        dt.Columns.Add(dc);
                    }
                }
                else
                {
                    aryLine = strLine.Split(',');
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < columnCount; j++)
                    {
                        dr[j] = aryLine[j];
                    }
                    dt.Rows.Add(dr);
                }
            }
            if (aryLine != null && aryLine.Length > 0)
            {
                dt.DefaultView.Sort = tableHead[0] + " " + "asc";
            }

            sr.Close();
            fs.Close();
            return dt;
        }
    }

    public class OpAD
    {
        //List<Dictionary<string, string>> ListUserInfo;
        Dictionary<string, string> UserInfo; //AD user account,email address
        string m_SamAccountName;

        public OpAD(string SamAccountName)
        {
            m_SamAccountName = SamAccountName;
        }

        //SamAccountName 为需要被查询email地址的用户账号，domainname为域名，username/password为登录域的账号密码，用来做认证
        static public string getEmailFromADUser(string SamAccountName, string domainname, string username,string password)
        {



            try
            {
                // enter AD settings  
                if(SamAccountName=="")
                {
                    return "NULL";
                }
                PrincipalContext AD = new PrincipalContext(ContextType.Domain, domainname, null, username, password);
                //PrincipalContext AD = new PrincipalContext(ContextType.Domain, "china.huawei.com");
                // create search user and add criteria  

                UserPrincipal u = new UserPrincipal(AD);
                u.SamAccountName = SamAccountName;

                // search for user  
                PrincipalSearcher search = new PrincipalSearcher(u);
                UserPrincipal result = (UserPrincipal)search.FindOne();
                search.Dispose();

                // show some details  


                Console.WriteLine("Display Name : " + result.DisplayName);
                Console.WriteLine("Email Address : " + result.EmailAddress);
                return result.EmailAddress;
            }

            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
                return "NULL";
            }



        }
        Dictionary<string, string>  getDetailsFromADUser(string SamAccountName)
        {



            try
            {
                // enter AD settings  
                PrincipalContext AD = new PrincipalContext(ContextType.Domain, "china.huawei.com");

                // create search user and add criteria  
                Console.WriteLine("Enter logon name: ");
                UserPrincipal u = new UserPrincipal(AD);
                u.SamAccountName = SamAccountName;

                // search for user  
                PrincipalSearcher search = new PrincipalSearcher(u);
                UserPrincipal result = (UserPrincipal)search.FindOne();
                search.Dispose();

                // show some details  
                UserInfo = new Dictionary<string, string>();

                UserInfo.Add("User Account", SamAccountName);
                UserInfo.Add("Display Name", result.DisplayName);
                UserInfo.Add("Email Address", result.EmailAddress);


                Console.WriteLine("Display Name : " + result.DisplayName);
                Console.WriteLine("Email Address : " + result.EmailAddress);
                return UserInfo;
            }

            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
                UserInfo = new Dictionary<string, string>();
                UserInfo.Add("User Account", SamAccountName);
                UserInfo.Add("Email Address", "Error: " + e.Message);
                return UserInfo;
            }

        }


    }

    /// <summary>
    /// Call way:
    ///   IniFile ini = new IniFile("C://test.ini");  
    ///   ini.IniWriteValue("LOC" ,"x" ,this.Location.X.ToString());
    ///   ini.IniWriteValue("LOC " ,"y" ,this.Location.Y.ToString());
    ///    
    /// </summary>
    public class IniFile
    {
        public string path;             //INI文件名  
        

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key,
                    string val, string filePath);

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def,
                    StringBuilder retVal, int size, string filePath);

        //声明读写INI文件的API函数  
        public IniFile(string INIPath)
        {
            path = INIPath;
        }

        //类的构造函数，传递INI文件名  
        public void IniWriteValue(string Section, string Key, string Value)
        {
            WritePrivateProfileString(Section, Key, Value, this.path);
        }

        //写INI文件  
        public string IniReadValue(string Section, string Key)
        {
            StringBuilder temp = new StringBuilder(255);
            int i = GetPrivateProfileString(Section, Key, "", temp, 255, this.path);
            return temp.ToString();
        }

        //读取INI文件指定  
    }
}
