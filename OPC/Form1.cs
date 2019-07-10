using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OPCAutomation;

namespace OPC
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OPCGroup OpcIn;//声明一个OPCGroup
        OPCServer OpcSvr = new OPCServer();//声明一个OPCserver
        private void Form1_Load(object sender, EventArgs e)
        {

            //注： 在我们新项目里面写OPC方法的时候 要先添加一个引用，在本程序的Debug 有，名字是“Interop.OPCAutomation.dll”    引用以后 上面在引用一下  using OPCAutomation;


            initopc();//窗体加载的时候调用OPC连接方法
        }
        void initopc()
        {
            OpcSvr.Connect("KEPware.KEPServerEx.V6", "127.0.0.1");//设置要连接OPC的版本和IP
            OpcIn = OpcSvr.OPCGroups.Add("OPCIN");//设置监控的OPCGroup

            OpcIn.OPCItems.AddItem("TEST.Device1.A01#_String", 1);//TEST.Device1.A01#_String  这个是KEPserver 里面添加的点   1 是 给这个监控点设置一个唯一的ID，
            OpcIn.OPCItems.AddItem("TEST.Device1.A02#_String", 2);
            OpcIn.OPCItems.AddItem("TEST.Device1.A03#_String", 3);
            OpcIn.OPCItems.AddItem("TEST.Device1.A04#_String", 4);

            OpcIn.OPCItems.AddItem("TEST.Device1.Pulsating1", 5);
            OpcIn.OPCItems.AddItem("TEST.Device1.Pulsating2", 6);
            OpcIn.OPCItems.AddItem("TEST.Device1.Pulsating3", 7);
            OpcIn.OPCItems.AddItem("TEST.Device1.Pulsating4", 8);
            OpcIn.OPCItems.AddItem("TEST.Device1.Pulsating5", 9);
            OpcIn.OPCItems.AddItem("TEST.Device1.TEST_1", 10);
            OpcIn.OPCItems.AddItem("TEST.Device1.TEST_2", 11);
   

            OpcIn.IsActive = true;//这个是标配
            OpcIn.IsSubscribed = true;//这个是标配
            OpcIn.DataChange += new DIOPCGroupEvent_DataChangeEventHandler(OpcInTri_DataChange);//给这个OPC采集的方法设置一个触发的方法，当OPC有改变的时候就会触发这个方法  
        }

         //OPC触发方法，这里就是取值
        void OpcInTri_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {   
            try
            {
                for (int i = 1; i <= NumItems; i++)//设置一个循环来取值
                {
                    string value = ItemValues.GetValue(i).ToString().ToLower().Trim();//value值   这个就是OPC的里面的 值 
                    int clientHandles = int.Parse(ClientHandles.GetValue(i).ToString());//唯一建   这个就是 上面添加点的时候注册的唯一ID
                    string tag = OpcIn.OPCItems.Item(clientHandles).ItemID;   //TAG点   这个点是KEP server里面我们添加的 点

                   //  然后 这里可以调用 我们写的 把数据存入数据库的方法
                   //  value 就是值
                    MessageBox.Show(tag+"点触发了"+" OPC取到的值是"+value);
                }
            }
            catch (Exception ex)
            {
            }
        }


       
        /// <summary>
        /// OPC写方法
        /// </summary>
        /// <param name="opcGroup">OPC组名</param>
        /// <param name="opcItem">OPC点名</param>
        /// <param name="value">待写入的值</param>
        private void WriteOpc(OPCGroup opcGroup, string opcItem, string value)
        {
            try
            {
                opcGroup.OPCItems.Item(opcItem).Write(value);
            }
            catch (Exception ex)
            {
                throw new Exception("写入OPC失败(WriteOpc):opcItem=" + opcItem + ",value=" + value + "  :" + ex.Message);
            }
        }

        /// <summary>
        /// 读取OPC的值
        /// </summary>
        /// <param name="opcGroup">OPC组名</param>
        /// <param name="opcItem">OPC点名</param>
        /// <returns>OPC的值</returns>
        private string ReadOpc(OPCGroup opcGroup, string opcItem)
        {
            string result = "";
            try
            {
                object value, quality, timestamp;
                opcGroup.OPCItems.Item(opcItem).Read(2, out value, out quality, out timestamp);

                if (value != null)
                {
                    result = value.ToString();
                }
                else
                {
                    result = "";
                }
            }
            catch (Exception ex)
            {
                throw new Exception("读取OPC失败(ReadOpc):opcItem=" + opcItem + "  :" + ex.Message);
            }
            return result;
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //  OpcIn 是OPCGroup  表示你要读取哪个OPCGroup 里面的哪个点
            string value = ReadOpc(OpcIn, textBox1.Text.ToString());//调用OPC读取单个地址的方法
            label1.Text = value;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // OpcIn 是OPCGroup  表示你要读写哪个OPCGroup 里面的哪个点
            WriteOpc(OpcIn,textBox2.Text.ToString(),textBox3.Text.ToString());//调用OPC写点的方法
        }

    }
}
